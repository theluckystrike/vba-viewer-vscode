'use strict';
// The three Zovo Pro tools: Export all modules, Macro security scan, Compare macros. Each asks pro.allowed() before
// it starts and calls pro.produced() only after its result is on disk or on screen (files written and read back, the
// scan table shown, the comparison table shown), so Escape, Cancel, an empty file or an error never uses a free try.
// Everything runs on this computer, and no macro is ever run.
const vscode = require('vscode');
const path = require('path');
const { loadFile, moduleUri, OFFICE_RE, isProjectBin, viewTypeFor } = require('./editor');
const { panelHtml } = require('./html');
const support = require('./support');
const S = require('./core/scan');
const X = require('./core/export');
const D = require('./core/diff');

const PREFIX = 'vbaViewer';
const setting = (k, d) => vscode.workspace.getConfiguration(PREFIX).get(k, d);
const base = (uri) => path.posix.basename(uri.path);
const stem = (name) => name.replace(/\.[^.]+$/, '');
const isTarget = (uri) => !!uri && (OFFICE_RE.test(uri.path) || isProjectBin(uri));
const OPEN_FIRST = 'Open the macros of an Office file first (right-click an .xlsm, .xls, .docm, .doc, .pptm or .ppt file and choose Open VBA Macros), then run this tool.';

function createTools({ context, provider, pro, reach }) {
  const panels = {};
  let last = {};

  async function exists(uri) { try { await vscode.workspace.fs.stat(uri); return true; } catch (e) { return false; } }
  async function readBytes(uri) { const d = await vscode.workspace.fs.readFile(uri); return Buffer.from(d.buffer, d.byteOffset, d.byteLength); }
  async function uniqueFile(dir, name) {
    let u = vscode.Uri.joinPath(dir, name);
    for (let i = 2; await exists(u) && i < 1000; i++) u = vscode.Uri.joinPath(dir, `${stem(name)}-${i}${path.posix.extname(name)}`);
    return u;
  }

  /** The file a tool should use: the clicked one, the one named by a viewer message, else the active viewer. */
  function targetUri(arg) {
    if (arg && arg.scheme) return arg;
    if (arg && typeof arg.uri === 'string') return vscode.Uri.parse(arg.uri);
    if (provider.active) return provider.active.uri;
    return undefined;
  }

  // ---------- panels ----------
  function openPanel(kind, title, onMessage) {
    const media = vscode.Uri.joinPath(context.extensionUri, 'media');
    let p = panels[kind];
    if (p && !p.disposed) { p.panel.title = title; p.onMessage = onMessage; p.panel.reveal(vscode.ViewColumn.Beside, false); return p; }
    const panel = vscode.window.createWebviewPanel(`${PREFIX}.${kind}`, title, { viewColumn: vscode.ViewColumn.Beside, preserveFocus: false }, {
      enableScripts: true, retainContextWhenHidden: true, localResourceRoots: [media],
      enableCommandUris: (reach ? reach.commandIds() : []).concat(support.commandUris(PREFIX)),
    });
    panel.webview.html = panelHtml({
      cspSource: panel.webview.cspSource,
      scriptUri: panel.webview.asWebviewUri(vscode.Uri.joinPath(media, 'panel.js')).toString(),
      styleUri: panel.webview.asWebviewUri(vscode.Uri.joinPath(media, 'viewer.css')).toString(),
      title, showSupportLinks: setting('showSupportLinks', true) !== false,
      head: reach ? (n) => reach.headHtml(n) : undefined,
    });
    let resolveReady;
    p = { panel, disposed: false, onMessage, ready: new Promise((r) => { resolveReady = r; }) };
    if (reach) reach.attach(panel.webview, panel);
    panel.onDidDispose(() => { p.disposed = true; });
    if (panel.onDidChangeViewState) panel.onDidChangeViewState(() => { if (panel.visible && reach) reach.refresh(); });
    panel.webview.onDidReceiveMessage(async (m) => {
      if (reach && reach.onMessage(m)) return;
      if (m && m.type === 'ready') { resolveReady(); return; }
      try { if (p.onMessage) await p.onMessage(m); } catch (e) { post(p, { type: 'error', message: String((e && e.message) || e) }); }
    });
    panels[kind] = p;
    return p;
  }
  async function post(p, msg) { await p.ready; try { await p.panel.webview.postMessage(msg); } catch (e) { /* closed */ } }

  async function loadOrSay(uri) {
    const r = await loadFile(uri);
    if (!r.ok) { vscode.window.showErrorMessage(`${base(uri)}: ${r.error}`); return null; }
    return r;
  }

  // ---------- Export all modules ----------
  async function exportTool(arg) {
    const uri = targetUri(arg);
    if (!isTarget(uri)) { vscode.window.showInformationMessage(OPEN_FIRST); return undefined; }
    if (!(await pro.allowed('exportModules', arg))) return undefined;
    const r = await loadOrSay(uri);
    if (!r) return undefined;
    if (!r.modules.length) { vscode.window.showInformationMessage(`${base(uri)} has no VBA modules, so there is nothing to export.`); return undefined; }
    const dir = vscode.Uri.joinPath(uri, '..');
    const defDir = vscode.Uri.joinPath(dir, `${stem(base(uri))}-vba`);
    const again = await exists(vscode.Uri.joinPath(defDir, X.MANIFEST));
    const pick = await vscode.window.showQuickPick([
      { label: again ? `$(sync) Re-export to ${base(defDir)}` : `$(folder) Export to ${base(defDir)}`, detail: again ? 'Only the modules that changed are rewritten' : `A new folder next to ${base(uri)}`, id: 'default' },
      { label: '$(folder-opened) Choose a folder...', id: 'browse' },
    ], { title: `Export all modules: ${base(uri)} (${r.modules.length} modules)`, placeHolder: 'Where should the .bas, .cls and .frm files go?', ignoreFocusOut: true });
    if (!pick) return undefined;
    let target = defDir;
    if (pick.id === 'browse') {
      const sel = await vscode.window.showOpenDialog({ canSelectFolders: true, canSelectFiles: false, canSelectMany: false, defaultUri: dir, openLabel: 'Export here' });
      if (!sel || !sel[0]) return undefined;
      target = sel[0];
    }
    await vscode.workspace.fs.createDirectory(target);
    const { files, encodingUsed, note } = X.exportFiles(r, setting('exportEncoding', 'utf8'));
    const existing = new Map();
    let oldManifest = null;
    try { oldManifest = JSON.parse((await readBytes(vscode.Uri.joinPath(target, X.MANIFEST))).toString('utf8')); } catch (e) { oldManifest = null; }
    const names = new Set(files.map((f) => f.file));
    if (oldManifest && Array.isArray(oldManifest.modules)) for (const m of oldManifest.modules) if (m && typeof m.file === 'string') names.add(m.file);
    for (const n of names) { const u = vscode.Uri.joinPath(target, X.safeName(n)); if (await exists(u)) existing.set(n.toLowerCase(), await readBytes(u)); }
    const plan = X.planExport(files, existing, oldManifest);
    for (const f of plan.write) {
      const u = vscode.Uri.joinPath(target, f.file);
      await vscode.workspace.fs.writeFile(u, f.bytes);
      const back = await readBytes(u);
      if (Buffer.compare(back, f.bytes) !== 0) throw new Error(`${f.file} was written, but reading it back gave other bytes.`);
    }
    await vscode.workspace.fs.writeFile(vscode.Uri.joinPath(target, X.MANIFEST), Buffer.from(X.manifest(base(uri), files, encodingUsed, new Date().toISOString()), 'utf8'));
    const writeSet = new Set(plan.write.map((f) => f.file));
    const rows = files.map((f) => ({ file: f.file, module: f.module, type: f.type, bytes: f.bytes.length, status: writeSet.has(f.file) ? (plan.write.find((w) => w.file === f.file).isNew ? 'new' : 'changed') : 'unchanged' }));
    const summary = `${plan.write.length} file${plan.write.length === 1 ? '' : 's'} written, ${plan.same.length} unchanged, in ${target.fsPath}.`;
    const p = openPanel('export', `Export: ${base(uri)}`, async (m) => {
      if (m && m.type === 'openFile' && typeof m.file === 'string' && names.has(m.file)) await vscode.window.showTextDocument(vscode.Uri.joinPath(target, m.file), { preview: false });
      else if (m && m.type === 'reveal') await vscode.commands.executeCommand('revealFileInOS', vscode.Uri.joinPath(target, X.MANIFEST));
    });
    await post(p, { type: 'export', file: base(uri), folder: target.fsPath, rows, summary, orphans: plan.orphans, encoding: encodingUsed, note });
    vscode.window.setStatusBarMessage(`Export all modules: ${summary}`, 10000);
    last = { ...last, exportModules: { folder: target.toString(), written: plan.write.map((f) => f.file), same: plan.same.map((f) => f.file), orphans: plan.orphans } };
    await pro.produced('exportModules');
    return last.exportModules;
  }

  // ---------- Macro security scan ----------
  async function scanTool(arg) {
    const uri = targetUri(arg);
    if (!isTarget(uri)) { vscode.window.showInformationMessage(OPEN_FIRST); return undefined; }
    if (!(await pro.allowed('scan', arg))) return undefined;
    const r = await loadOrSay(uri);
    if (!r) return undefined;
    if (!r.modules.length && !r.xlm.length) { vscode.window.showInformationMessage(`${base(uri)} has no VBA modules and no Excel 4.0 macro sheets, so there is nothing to scan.`); return undefined; }
    const rep = S.scan(r, { fileName: base(uri) });
    const date = new Date().toISOString().replace('T', ' ').slice(0, 16);
    const p = openPanel('scan', `Scan: ${base(uri)}`, async (m) => {
      if (!m || m.type !== 'save' || !/^(md|json)$/.test(m.format)) return;
      const dir = vscode.Uri.joinPath(uri, '..');
      const out = await uniqueFile(dir, `${stem(base(uri))}-macro-scan.${m.format}`);
      const text = m.format === 'md' ? S.toMarkdown(rep, { file: base(uri), date }) : `${JSON.stringify({ file: base(uri), scanned: date, tool: 'VBA Viewer & Macro Extractor (Zovo)', ...rep }, null, 2)}\n`;
      await vscode.workspace.fs.writeFile(out, Buffer.from(text, 'utf8'));
      const back = (await readBytes(out)).toString('utf8');
      if (back !== text) throw new Error(`${base(out)} was written, but reading it back gave other text.`);
      post(p, { type: 'saved', file: base(out), path: out.fsPath });
      last = { ...last, report: out.toString() };
      await vscode.window.showTextDocument(out, { preview: false, viewColumn: vscode.ViewColumn.Beside });
    });
    await post(p, { type: 'scan', file: base(uri), report: rep });
    last = { ...last, scan: rep };
    await pro.produced('scan');
    return rep;
  }

  // ---------- Compare macros ----------
  async function compareTool(arg) {
    const uri = targetUri(arg);
    if (!isTarget(uri)) { vscode.window.showInformationMessage(OPEN_FIRST); return undefined; }
    if (!(await pro.allowed('compare', arg))) return undefined;
    const a = await loadOrSay(uri);
    if (!a) return undefined;
    const found = (await vscode.workspace.findFiles('**/*.{xlsm,xlsb,xls,xltm,xlam,docm,dotm,doc,pptm,ppt,bin}', '**/node_modules/**', 2000))
      .filter((u) => u.toString() !== uri.toString() && isTarget(u));
    const items = found.map((u) => ({ label: base(u), description: vscode.workspace.asRelativePath(u), u }));
    items.push({ label: '$(folder-opened) Choose a file...', id: 'browse', alwaysShow: true });
    const pick = await vscode.window.showQuickPick(items, { title: `Compare macros: ${base(uri)} with...`, placeHolder: 'Pick the other Office file', ignoreFocusOut: true, matchOnDescription: true });
    if (!pick) return undefined;
    let other = pick.u;
    if (pick.id === 'browse') {
      const sel = await vscode.window.showOpenDialog({ canSelectMany: false, openLabel: 'Compare', filters: { 'Office files with macros': ['xlsm', 'xlsb', 'xls', 'xltm', 'xlam', 'docm', 'dotm', 'doc', 'pptm', 'ppt', 'bin'] } });
      if (!sel || !sel[0]) return undefined;
      other = sel[0];
    }
    const b = await loadOrSay(other);
    if (!b) return undefined;
    const cmp = D.compareResults(a, b);
    const p = openPanel('compare', `Compare: ${base(uri)} / ${base(other)}`, async (m) => {
      if (!m || m.type !== 'diff' || typeof m.name !== 'string') return;
      const row = cmp.rows.find((x) => x.name === m.name);
      if (!row) return;
      if (row.status === 'onlyA') { await provider.openModuleText(uri, row.name, 1, vscode.ViewColumn.Active); return; }
      if (row.status === 'onlyB') { await provider.openModuleText(other, row.name, 1, vscode.ViewColumn.Active); return; }
      const left = moduleUri(uri, row.name);
      const right = moduleUri(other, row.name);
      await vscode.commands.executeCommand('vscode.diff', left, right, `${row.name}: ${base(uri)} ↔ ${base(other)}`, { preview: false });
      last = { ...last, diff: row.name };
    });
    await post(p, { type: 'compare', a: base(uri), b: base(other), rows: cmp.rows, sum: cmp.sum });
    last = { ...last, compare: { a: uri.toString(), b: other.toString(), ...cmp } };
    await pro.produced('compare');
    return last.compare;
  }

  return { exportTool, scanTool, compareTool, panels, get last() { return last; }, viewTypeFor };
}

module.exports = { createTools };
