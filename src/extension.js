'use strict';
// VBA Viewer & Macro Extractor by Zovo: shows the VBA macros inside Office files (.xlsm, .xlsb, .xls, .xltm, .xlam,
// .docm, .dotm, .doc, .pptm, .ppt and vbaProject.bin) in a read-only VS Code tab. Macros are read as text and never
// run; nothing is written to the file. No network, no telemetry. Activation shows nothing: no message, no popup.
const vscode = require('vscode');
const { VbaEditorProvider, PROJECT_VIEW, MACROS_VIEW, viewTypeFor, createModuleContentProvider, SCHEME, OFFICE_RE, isProjectBin } = require('./editor');
const { createTools } = require('./tools');
const { createPro } = require('./pro');
const { createRatingPrompt } = require('./rate');
const support = require('./support');
const { createReach } = require('./reach');
const reachConfig = require('./reach-config');

const PREFIX = 'vbaViewer';
const EXT_ID = 'zovo.zovo-vba-viewer';
const setting = (k, d) => vscode.workspace.getConfiguration(PREFIX).get(k, d);

function activate(context) {
  const sub = (d) => context.subscriptions.push(d);
  let rating = { success() {}, failure() {} };
  try { rating = createRatingPrompt({ vscode, context, extensionId: EXT_ID, section: PREFIX, name: 'VBA Viewer' }); } catch (e) { /* rating is optional */ }
  support.registerSupport({ vscode, context, prefix: PREFIX, selfId: EXT_ID });
  const pro = createPro({ vscode, context });
  const reach = createReach({ vscode, ...reachConfig, context, licensed: () => !!pro.licensed, shown: () => setting('showProTools', true) !== false });
  reach.registerCtx();
  sub(pro.onDidChange(() => reach.refresh()));

  const provider = new VbaEditorProvider(context, {
    reach: () => reach,
    onSuccess: (doc) => rating.success(doc.uri),
    onFailure: () => rating.failure(),
  });
  for (const vt of [PROJECT_VIEW, MACROS_VIEW]) sub(vscode.window.registerCustomEditorProvider(vt, provider, { webviewOptions: { retainContextWhenHidden: true }, supportsMultipleEditorsPerDocument: true }));
  sub(vscode.workspace.registerTextDocumentContentProvider(SCHEME, createModuleContentProvider()));
  const tools = createTools({ context, provider, pro, reach });

  const reg = (id, fn) => sub(vscode.commands.registerCommand(id, async (...a) => {
    try { return await fn(...a); } catch (e) { vscode.window.showErrorMessage(String((e && e.message) || e)); return undefined; }
  }));
  const target = (uri) => (uri && (uri.scheme || typeof uri.uri === 'string') ? uri : undefined);
  reg(`${PREFIX}.openMacros`, (uri) => {
    let u = uri && uri.scheme ? uri : undefined;
    if (!u) {
      const tab = vscode.window.tabGroups && vscode.window.tabGroups.activeTabGroup && vscode.window.tabGroups.activeTabGroup.activeTab;
      const input = tab && tab.input;
      if (input && input.uri && input.uri.scheme) u = input.uri;
    }
    if (!u && vscode.window.activeTextEditor) u = vscode.window.activeTextEditor.document.uri;
    if (!u || !(OFFICE_RE.test(u.path) || isProjectBin(u))) { vscode.window.showInformationMessage('Select an .xlsm, .xlsb, .xls, .xltm, .xlam, .docm, .dotm, .doc, .pptm, .ppt or vbaProject.bin file first.'); return undefined; }
    return vscode.commands.executeCommand('vscode.openWith', u, viewTypeFor(u));
  });
  reg(`${PREFIX}.openDefault`, (uri) => {
    const u = uri && uri.scheme ? uri : provider.active && provider.active.uri;
    if (!u) return undefined;
    return vscode.commands.executeCommand('vscode.openWith', u, 'default');
  });
  reg(`${PREFIX}.openSample`, () => vscode.commands.executeCommand('vscode.openWith', vscode.Uri.joinPath(context.extensionUri, 'samples', 'sample-macros.xlsm'), MACROS_VIEW));
  reg(`${PREFIX}.openWalkthrough`, () => vscode.commands.executeCommand('workbench.action.openWalkthrough', `${EXT_ID}#${PREFIX}.start`, false));
  reg(`${PREFIX}.exportModules`, reach.wrap((arg) => tools.exportTool(target(arg))));
  reg(`${PREFIX}.scan`, reach.wrap((arg) => tools.scanTool(target(arg))));
  reg(`${PREFIX}.compare`, reach.wrap((arg) => tools.compareTool(target(arg))));
  sub(vscode.workspace.onDidChangeConfiguration((e) => {
    if (e.affectsConfiguration(`${PREFIX}.showProTools`)) reach.refresh();
    if (e.affectsConfiguration(`${PREFIX}.showAttributes`) || e.affectsConfiguration(`${PREFIX}.showProTools`) || e.affectsConfiguration(`${PREFIX}.maxFileSizeMB`)) for (const d of provider.docs) provider.reload(d);
  }));
  sub(vscode.window.registerUriHandler({ handleUri: (uri) => pro.handleUri(uri) }));
  return { provider, tools, pro };
}

function deactivate() {}

module.exports = { activate, deactivate };
