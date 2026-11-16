'use strict';
// Read-only custom editor for the VBA macros of an Office file. "vbaViewer.project" opens vbaProject.bin by default;
// "vbaViewer.macros" is an optional editor for .xlsm, .xlsb, .xls, .xltm, .xlam, .docm, .dotm, .doc, .pptm and .ppt,
// so the user's spreadsheet, document and slide viewers stay the default for those files. The extension host reads
// the file (src/core) and sends plain data to the webview; nothing runs and nothing leaves the computer
// (CSP default-src 'none').
const vscode = require('vscode');
const path = require('path');
const { extract, TYPE_LABEL } = require('./core/vba');
const { editorHtml } = require('./html');
const support = require('./support');

const PROJECT_VIEW = 'vbaViewer.project';
const MACROS_VIEW = 'vbaViewer.macros';
const PREFIX = 'vbaViewer';
const OFFICE_RE = /\.(xlsm|xlsb|xls|xltm|xlam|docm|dotm|doc|pptm|ppt)$/i;
const setting = (k, d) => vscode.workspace.getConfiguration(PREFIX).get(k, d);
const isProjectBin = (uri) => /(^|\/)vbaproject\.bin$/i.test(uri.path);
const viewTypeFor = (uri) => (isProjectBin(uri) ? PROJECT_VIEW : MACROS_VIEW);

/** Reads and parses a file. Returns the extract() result plus { size, name } or { ok: false, error }. */
async function loadFile(uri) {
  const name = path.posix.basename(uri.path);
  try {
    const stat = await vscode.workspace.fs.stat(uri);
    const max = Math.max(1, Number(setting('maxFileSizeMB', 200)) || 200);
    if (stat.size > max * 1048576) return { ok: false, name, size: stat.size, error: `This file is ${(stat.size / 1048576).toFixed(0)} MB, more than the ${max} MB set in "vbaViewer.maxFileSizeMB". Raise the setting to read it.` };
    const d = await vscode.workspace.fs.readFile(uri);
    const r = extract(Buffer.from(d.buffer, d.byteOffset, d.byteLength), name);
    return { ...r, name, size: stat.size };
  } catch (e) {
    return { ok: false, name, error: `The file could not be read: ${String((e && e.message) || e)}` };
  }
}

class VbaDocument {
  constructor(uri) { this.uri = uri; this.result = null; this.panels = new Set(); this.listeners = []; }
  get name() { return path.posix.basename(this.uri.path); }
  async load() { this.result = await loadFile(this.uri); }
  dispose() { for (const l of this.listeners) l.dispose(); }
}

/** The message the webview gets for a document. */
function payload(doc) {
  const r = doc.result || { ok: false, error: 'The file could not be read.' };
  const base = { type: 'load', name: doc.name, size: r.size || 0, showPro: setting('showProTools', true) !== false, showAttributes: setting('showAttributes', false) === true };
  if (!r.ok) return { ...base, error: r.error };
  return {
    ...base,
    container: r.container, host: r.host, vbaPath: r.vbaPath, hasVba: r.hasVba, note: r.note, protected: r.protected, warnings: r.warnings,
    project: r.project,
    references: r.references,
    xlm: r.xlm, xlmNames: r.xlmNames,
    modules: r.modules.map((m) => ({ name: m.name, type: m.type, typeLabel: TYPE_LABEL[m.type] || m.type, streamName: m.streamName, source: m.source, size: m.size, compressedSize: m.compressedSize, offset: m.offset, lines: m.lines, attributes: m.attributes, procedures: m.procedures, error: m.error, readOnly: m.readOnly, private: m.private, docString: m.docString, hasFrame: !!m.frame })),
  };
}

class VbaEditorProvider {
  constructor(context, hooks = {}) {
    this.context = context;
    this.hooks = hooks;
    this.docs = new Set();
    this.active = null;
    this.selected = new Map();
    this._success = new WeakSet();
  }

  async openCustomDocument(uri) {
    const doc = new VbaDocument(uri);
    await doc.load();
    this.docs.add(doc);
    if (uri.scheme === 'file') {
      try {
        const w = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(vscode.Uri.joinPath(uri, '..'), path.posix.basename(uri.path)));
        let timer = null;
        const reload = () => { clearTimeout(timer); timer = setTimeout(() => this.reload(doc), 300); };
        doc.listeners.push(w, w.onDidChange(reload), w.onDidCreate(reload));
      } catch (e) { /* no watcher: the view still works */ }
    }
    return { uri, dispose: () => { this.docs.delete(doc); doc.dispose(); if (this.active === doc) this.active = null; }, _doc: doc };
  }

  async reload(doc) { await doc.load(); for (const p of doc.panels) this.send(p.webview, payload(doc)); }

  async resolveCustomEditor(handle, panel) {
    const doc = handle._doc;
    doc.panels.add(panel);
    const media = vscode.Uri.joinPath(this.context.extensionUri, 'media');
    const reach = this.hooks.reach ? this.hooks.reach() : null;
    panel.webview.options = {
      enableScripts: true, localResourceRoots: [media],
      enableCommandUris: (reach ? reach.commandIds() : []).concat(support.commandUris(PREFIX)),
    };
    panel.webview.html = editorHtml({
      cspSource: panel.webview.cspSource,
      scriptUri: panel.webview.asWebviewUri(vscode.Uri.joinPath(media, 'viewer.js')).toString(),
      styleUri: panel.webview.asWebviewUri(vscode.Uri.joinPath(media, 'viewer.css')).toString(),
      title: doc.name,
      head: reach ? (n) => reach.headHtml(n) : undefined,
      showSupportLinks: setting('showSupportLinks', true) !== false,
    });
    if (reach) reach.attach(panel.webview, panel);
    const setActive = () => {
      if (panel.active) this.active = doc;
      // A Pro tool that opened its result in another tab leaves this row stale when the user comes back.
      if (panel.visible && reach) reach.refresh();
    };
    setActive();
    panel.onDidChangeViewState(setActive);
    panel.onDidDispose(() => { doc.panels.delete(panel); if (this.active === doc && ![...doc.panels].length) this.active = null; });
    panel.webview.onDidReceiveMessage(async (m) => {
      if (reach && reach.onMessage(m)) return;
      if (!m || typeof m !== 'object') return;
      try {
        if (m.type === 'ready') this.send(panel.webview, payload(doc));
        else if (m.type === 'shown') { if (!this._success.has(doc)) { this._success.add(doc); if (this.hooks.onSuccess) this.hooks.onSuccess(doc); } }
        else if (m.type === 'failed') { if (this.hooks.onFailure) this.hooks.onFailure(doc); }
        else if (m.type === 'select' && typeof m.module === 'string') this.selected.set(doc, m.module);
        else if (m.type === 'copy' && typeof m.text === 'string') { await vscode.env.clipboard.writeText(m.text.slice(0, 5000000)); vscode.window.setStatusBarMessage(`Copied ${m.what || 'the source'}`, 3000); }
        else if (m.type === 'openText' && typeof m.module === 'string') await this.openModuleText(doc.uri, m.module, Math.max(1, m.line | 0));
        else if (m.type === 'tool' && /^(exportModules|scan|compare)$/.test(m.id)) await vscode.commands.executeCommand(`${PREFIX}.${m.id}`, { surface: 'pill', uri: doc.uri.toString() });
        else if (m.type === 'openDefault') await vscode.commands.executeCommand('vscode.openWith', doc.uri, 'default');
      } catch (e) { vscode.window.showErrorMessage(String((e && e.message) || e)); }
    });
  }

  /** Opens one module as read-only text (scheme vba-macro) with the Visual Basic language. */
  async openModuleText(fileUri, module, line = 1, column) {
    const uri = moduleUri(fileUri, module);
    const d = await vscode.workspace.openTextDocument(uri);
    try { if (d.languageId !== 'vb') await vscode.languages.setTextDocumentLanguage(d, 'vb'); } catch (e) { /* no VB language: plain text */ }
    const pos = new vscode.Position(Math.min(line - 1, Math.max(0, (d.lineCount || 1) - 1)), 0);
    return vscode.window.showTextDocument(d, { preview: false, selection: new vscode.Range(pos, pos), viewColumn: column || vscode.ViewColumn.Beside });
  }

  send(webview, msg) { try { const r = webview.postMessage(msg); if (r && r.then) r.then(undefined, () => false); } catch (e) { /* disposed */ } }
}

const SCHEME = 'vba-macro';
/** vba-macro:/<file name>/<module>.bas?<encoded file uri> */
function moduleUri(fileUri, module, ext = '.bas') {
  const name = path.posix.basename(fileUri.path);
  return vscode.Uri.from({ scheme: SCHEME, path: `/${name}/${module}${ext}`, query: encodeURIComponent(fileUri.toString()) });
}

/** Text documents of single modules for "Open as text" and the compare diff. Read-only (no file system provider). */
function createModuleContentProvider() {
  return {
    async provideTextDocumentContent(uri) {
      const src = vscode.Uri.parse(decodeURIComponent(uri.query));
      const r = await loadFile(src);
      if (!r.ok) return `' ${r.error}\r\n`;
      const want = path.posix.basename(uri.path).replace(/\.[^.]+$/, '');
      const m = r.modules.find((x) => x.name === want) || r.modules.find((x) => x.name.toLowerCase() === want.toLowerCase());
      if (!m) return `' The module "${want}" is not in ${r.name}.\r\n`;
      return m.source;
    },
  };
}

module.exports = { VbaEditorProvider, VbaDocument, PROJECT_VIEW, MACROS_VIEW, OFFICE_RE, viewTypeFor, payload, loadFile, moduleUri, createModuleContentProvider, SCHEME, isProjectBin };
