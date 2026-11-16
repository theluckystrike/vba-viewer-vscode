'use strict';
// Extension tests with the vscode stub: listing, activation (nothing shows), the viewer's messages, the three Pro tools
// with 1 free try each (file A runs, then file B opens the Pro panel; a reload keeps the count), README, CSP, support.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');

process.env.ZOVO_PRO_FILE = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'vba-pro-')), 'pro.json');
const { loadExtension, tmpDir, cleanup } = require('./support/vscode-stub');
const { editorHtml, panelHtml } = require('../src/html');

const ROOT = path.join(__dirname, '..');
const FIX = path.join(__dirname, 'fixtures');
const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
const tick = (ms = 5) => new Promise((r) => setTimeout(r, ms));
const until = async (f, ms = 3000) => { const t = Date.now(); while (Date.now() - t < ms) { const v = await f(); if (v) return v; await tick(10); } return f(); };
test.after(() => { cleanup(); fs.rmSync(path.dirname(process.env.ZOVO_PRO_FILE), { recursive: true, force: true }); });

function activate(folder, state) {
  const s = loadExtension();
  if (state) for (const [k, v] of state) s.state.set(k, JSON.parse(JSON.stringify(v)));
  if (folder) s.vscode.workspace.workspaceFolders = [{ uri: s.vscode.Uri.file(folder), name: 'ws' }];
  s.api = s.ext.activate(s.context);
  return s;
}
function tmpWs() {
  const dir = tmpDir('vba-ws-');
  fs.mkdirSync(path.join(dir, 'books'), { recursive: true });
  fs.copyFileSync(path.join(ROOT, 'samples', 'sample-macros.xlsm'), path.join(dir, 'books', 'budget.xlsm'));
  fs.copyFileSync(path.join(FIX, 'sample-macros-v2.xlsm'), path.join(dir, 'books', 'budget-v2.xlsm'));
  fs.copyFileSync(path.join(FIX, 'real', 'SimpleMacro.xls'), path.join(dir, 'books', 'SimpleMacro.xls'));
  fs.copyFileSync(path.join(FIX, 'real', 'SimpleMacro.docm'), path.join(dir, 'books', 'SimpleMacro.docm'));
  fs.copyFileSync(path.join(FIX, 'no-macros.xlsm'), path.join(dir, 'books', 'no-macros.xlsm'));
  fs.copyFileSync(path.join(FIX, 'broken', 'not-office.xls'), path.join(dir, 'books', 'not-office.xls'));
  return dir;
}
function fakePanel() {
  let onMsg;
  const posted = [];
  return {
    posted, viewColumn: 1, active: true, visible: true,
    webview: { options: {}, html: '', cspSource: 'vscode-resource:', asWebviewUri: (u) => u, postMessage: async (m) => { posted.push(m); return true; }, onDidReceiveMessage: (f) => { onMsg = f; return { dispose() {} }; } },
    onDidChangeViewState: () => ({ dispose() {} }), onDidDispose: () => ({ dispose() {} }), reveal() {},
    send: (m) => onMsg(m),
  };
}
async function openEditor(s, file) {
  const prov = s.providers.custom.p;
  const handle = await prov.openCustomDocument(s.vscode.Uri.file(file));
  const panel = fakePanel();
  await prov.resolveCustomEditor(handle, panel);
  await panel.send({ type: 'ready' });
  return { prov, handle, panel, doc: handle._doc, load: panel.posted.find((m) => m.type === 'load') };
}
const proPanel = (s, tool) => s.log.info.find((m) => m.opts && m.opts.modal && m.text === `${tool} is part of Zovo Pro.`);
const uriOf = (s, dir, f) => s.vscode.Uri.file(path.join(dir, 'books', f));

// ---------- package.json and listing ----------
test('listing: display name with the phrase first, 30 characters or less, the row keywords, engines for Cursor and Windsurf', () => {
  assert.equal(pkg.displayName, 'VBA Viewer & Macro Extractor');
  assert.ok(pkg.displayName.length <= 30);
  assert.ok(pkg.description.startsWith('VBA viewer'));
  assert.deepEqual(pkg.keywords, ['vba', 'excel macro', 'xlsm', 'macro', 'office']);
  assert.equal(pkg.engines.vscode, '^1.75.0');
  assert.equal(`${pkg.publisher}.${pkg.name}`, 'zovo.zovo-vba-viewer');
  assert.equal(pkg.version, '0.1.0');
  assert.equal(pkg.license, 'MIT');
  assert.equal(pkg.repository.url, 'https://github.com/theluckystrike/vba-viewer-vscode');
  assert.deepEqual(pkg.activationEvents, ['onUri']);
});
test('listing: vbaProject.bin is the only default; Office files get an optional editor and Open VBA Macros in the Explorer and editor title', () => {
  const ed = Object.fromEntries(pkg.contributes.customEditors.map((e) => [e.viewType, e]));
  assert.equal(ed['vbaViewer.project'].priority, 'default');
  assert.deepEqual(ed['vbaViewer.project'].selector.map((s) => s.filenamePattern), ['vbaProject.bin']);
  assert.equal(ed['vbaViewer.macros'].priority, 'option');
  assert.deepEqual(ed['vbaViewer.macros'].selector.map((s) => s.filenamePattern), ['*.xlsm', '*.xlsb', '*.xls', '*.xltm', '*.xlam', '*.docm', '*.dotm', '*.doc', '*.pptm', '*.ppt']);
  const ctx = pkg.contributes.menus['explorer/context'];
  for (const m of [...ctx, ...pkg.contributes.menus['editor/title']]) { assert.equal(m.title, undefined); assert.ok(pkg.contributes.commands.some((c) => c.command === m.command)); }
  assert.equal(ctx[0].command, 'vbaViewer.openMacros');
  const re = new RegExp(ctx[0].when.match(/resourceExtname =~ \/(.*?)\/i/)[1], 'i');
  for (const f of ['.xlsm', '.XLSM', '.xls', '.docm', '.DOC', '.ppt', '.pptm', '.xlam', '.xltm', '.dotm', '.xlsb']) assert.ok(re.test(f), f);
  for (const f of ['.xlsx', '.docx', '.pptx', '.csv', '.bin']) assert.ok(!re.test(f), f);
  assert.match(ctx[0].when, /resourceFilename =~ \/\^vbaproject\\\.bin\$\/i/);
  assert.ok(!/\\\\\\\\/.test(JSON.stringify(pkg.contributes.menus)), 'no doubled escaping in the menu regexes');
  assert.ok(pkg.contributes.menus['editor/title'].some((m) => m.command === 'vbaViewer.openMacros' && /activeCustomEditorId != 'vbaViewer.macros'/.test(m.when)));
  const titles = Object.fromEntries(pkg.contributes.commands.map((c) => [c.command, c.title]));
  assert.equal(titles['vbaViewer.openMacros'], 'Open VBA Macros');
  for (const id of ['exportModules', 'scan', 'compare']) {
    assert.match(titles[`vbaViewer.zovoPro.${id}`], /^Zovo Pro: /);
    assert.ok(pkg.contributes.menus.commandPalette.some((m) => m.command === `vbaViewer.zovoPro.${id}` && m.when === 'false'));
    assert.ok(ctx.some((m) => m.command === `vbaViewer.zovoPro.${id}`));
  }
  assert.ok(fs.existsSync(path.join(ROOT, 'icon.png')));
  const png = fs.readFileSync(path.join(ROOT, 'icon.png'));
  assert.deepEqual([png.readUInt32BE(16), png.readUInt32BE(20)], [128, 128]);
});
test('pro.js keeps the free tries in one constant (1); proinfo has no triesMax; tries key and tools match', () => {
  const pro = require('../src/pro');
  assert.equal(pro.TRIAL_USES, 1);
  assert.equal(pro.TRIAL_KEY, 'vbaViewer.proTrial');
  const cfg = require('../src/license/config');
  assert.equal(cfg.listing, 'vba');
  assert.equal(cfg.pro_info.triesKey, pro.TRIAL_KEY);
  assert.equal(cfg.pro_info.triesMax, undefined);
  assert.deepEqual(cfg.pro_info.tools.map((t) => t.id), Object.keys(pro.FEATURES));
  assert.deepEqual(require('../src/reach-config').tools.map((t) => t.id), Object.keys(pro.FEATURES));
  assert.equal(require('../src/license').MODULE_VERSION, '2026-11-15.tries1');
});

// ---------- activation ----------
test('activation shows nothing: no message, no panel, no editor, no browser; every command is registered', () => {
  const s = activate();
  assert.equal(s.log.info.length + s.log.warn.length + s.log.error.length, 0);
  assert.equal(s.log.panel, undefined);
  assert.equal(s.log.opened.length, 0);
  assert.equal(s.log.external.length, 0);
  for (const c of pkg.contributes.commands) assert.ok(s.commands.has(c.command), `registered: ${c.command}`);
});

// ---------- viewer ----------
test('viewer: the sample opens with its modules, sources, project and nothing to fix', async () => {
  const s = activate();
  const { load } = await openEditor(s, path.join(ROOT, 'samples', 'sample-macros.xlsm'));
  assert.equal(load.hasVba, true);
  assert.equal(load.container, 'ooxml');
  assert.equal(load.host, 'Excel');
  assert.equal(load.project.name, 'BudgetMacros');
  assert.deepEqual(load.modules.map((m) => m.typeLabel), ['Document module', 'Document module', 'Standard module', 'Standard module', 'Class module']);
  assert.match(load.modules.find((m) => m.name === 'BudgetReport').source, /Public Sub BuildSummary\(\)/);
  assert.equal(load.showAttributes, false);
  assert.equal(load.showPro, true);
});
test('viewer: no-VBA file gets the empty state, a broken file a plain error; the file is never written', async () => {
  const dir = tmpWs();
  const s = activate(dir);
  const before = fs.readFileSync(path.join(dir, 'books', 'no-macros.xlsm'));
  const n = await openEditor(s, path.join(dir, 'books', 'no-macros.xlsm'));
  assert.equal(n.load.hasVba, false);
  assert.match(n.load.note, /no vbaProject\.bin part/);
  assert.equal(n.load.error, undefined);
  const b = await openEditor(s, path.join(dir, 'books', 'not-office.xls'));
  assert.match(b.load.error, /plain text/);
  assert.deepEqual(fs.readFileSync(path.join(dir, 'books', 'no-macros.xlsm')), before);
  s.config.maxFileSizeMB = 1;
  fs.writeFileSync(path.join(dir, 'books', 'big.xlsm'), Buffer.alloc(1100000));
  const big = await openEditor(s, path.join(dir, 'books', 'big.xlsm'));
  assert.match(big.load.error, /more than the .* MB set in "vbaViewer\.maxFileSizeMB"/);
});
test('viewer messages: copy, open a module as text (vba-macro document with the source), the Pro pill passes its surface', async () => {
  const dir = tmpWs();
  const s = activate(dir);
  const file = path.join(dir, 'books', 'budget.xlsm');
  const { panel } = await openEditor(s, file);
  await panel.send({ type: 'copy', text: 'Sub A()', what: 'A' });
  assert.equal(s.log.clipboard, 'Sub A()');
  await panel.send({ type: 'openText', module: 'BudgetReport', line: 10 });
  assert.equal(s.log.opened.length, 1);
  const d = s.log.opened[0];
  assert.equal(d.uri.scheme, 'vba-macro');
  assert.match(d.getText(), /^Attribute VB_Name = "BudgetReport"/);
  assert.equal(d.languageId, 'vb');
  s.state.set('vbaViewer.proTrial', { scan: 1 });
  await panel.send({ type: 'tool', id: 'scan' });
  const call = s.log.commands.find((c) => c[0] === 'vbaViewer.scan');
  assert.equal(call[1].surface, 'pill');
  assert.equal(call[1].uri, s.vscode.Uri.file(file).toString());
  assert.ok(proPanel(s, 'Macro security scan'), 'no try left: the Pro panel after the click');
});
test('commands: Open VBA Macros opens the optional editor for Office files and the project editor for vbaProject.bin', async () => {
  const s = activate();
  await s.commands.get('vbaViewer.openMacros')(s.vscode.Uri.file('/x/Book1.XLSM'));
  await s.commands.get('vbaViewer.openMacros')(s.vscode.Uri.file('/x/vbaProject.bin'));
  const calls = s.log.commands.filter((c) => c[0] === 'vscode.openWith').map((c) => c[2]);
  assert.deepEqual(calls, ['vbaViewer.macros', 'vbaViewer.project']);
  await s.commands.get('vbaViewer.openMacros')(s.vscode.Uri.file('/x/notes.txt'));
  assert.match(s.log.info[0].text, /^Select an \.xlsm/);
  await s.commands.get('vbaViewer.openSample')();
  assert.match(s.log.commands.at(-1)[1].fsPath, /samples\/sample-macros\.xlsm$/);
});
test('html: CSP allows no remote origin, no inline style, scripts by nonce only; media makes no network calls', () => {
  const h = editorHtml({ cspSource: 'vscode-resource:', scriptUri: 'a.js', styleUri: 'a.css', title: 't', nonce: 'N' });
  const csp = h.match(/content="([^"]+)"/)[1];
  assert.match(csp, /^default-src 'none'/);
  assert.doesNotMatch(csp, /https?:|connect-src|unsafe-inline|unsafe-eval/);
  assert.match(csp, /script-src 'nonce-N'/);
  assert.match(h, /id="toolbar"/);
  assert.match(h, /never run/);
  const pnl = panelHtml({ cspSource: 'vscode-resource:', scriptUri: 'p.js', styleUri: 'a.css', title: 't', nonce: 'M' });
  assert.match(pnl, /script-src 'nonce-M'/);
  for (const f of ['viewer.js', 'panel.js']) {
    const js = fs.readFileSync(path.join(ROOT, 'media', f), 'utf8');
    assert.doesNotMatch(js, / style="/, `${f}: inline style attributes break the CSP`);
    assert.doesNotMatch(js, /\bfetch\(|XMLHttpRequest|WebSocket|https?:\/\//, `${f}: no network`);
  }
});

// ---------- Pro tools ----------
test('Pro export: file A writes .bas/.cls files and counts the try; file B then opens the Pro panel and writes nothing', async () => {
  const dir = tmpWs();
  const s = activate(dir);
  const a = uriOf(s, dir, 'budget.xlsm');
  assert.equal(s.api.pro.left('exportModules'), 1);
  s.ui.quick.push(undefined); // Escape at the folder pick: no try used
  assert.equal(await s.api.tools.exportTool(a), undefined);
  assert.equal(s.api.pro.left('exportModules'), 1);
  s.ui.quick.push((items) => items.find((i) => i.id === 'default'));
  const r = await s.api.tools.exportTool(a);
  assert.equal(s.log.lastPick.opts.ignoreFocusOut, true);
  const out = path.join(dir, 'books', 'budget-vba');
  assert.deepEqual(fs.readdirSync(out).sort(), ['.vba-export.json', 'BudgetReport.bas', 'Formatting.bas', 'LineItem.cls', 'Sheet1.cls', 'ThisWorkbook.cls']);
  assert.match(fs.readFileSync(path.join(out, 'LineItem.cls'), 'utf8'), /^VERSION 1\.0 CLASS\r\n[\s\S]*Public Property Let Amount\(ByVal value As Currency\)/);
  assert.equal(r.written.length, 5);
  const posted = s.log.panel.posted.find((m) => m.type === 'export');
  assert.equal(posted.rows.length, 5);
  assert.match(posted.summary, /^5 files written, 0 unchanged/);
  assert.equal(s.api.pro.left('exportModules'), 0);
  assert.match(s.log.status.join('\n'), /Zovo Pro: the free try of "Export all modules" is used on this computer\./);
  // file B
  const n = s.log.posted.length;
  assert.equal(await s.api.tools.exportTool(uriOf(s, dir, 'SimpleMacro.docm')), undefined);
  assert.ok(proPanel(s, 'Export all modules'));
  assert.equal(s.log.posted.length, n);
  assert.ok(!fs.existsSync(path.join(dir, 'books', 'SimpleMacro-vba')));
  assert.equal(s.api.pro.left('scan'), 1);
  assert.equal(s.api.pro.left('compare'), 1);
});
test('Pro export: a re-export (new profile) rewrites only the changed module and keeps the others untouched', async () => {
  const dir = tmpWs();
  const s1 = activate(dir);
  const a = uriOf(s1, dir, 'budget.xlsm');
  s1.ui.quick.push((items) => items.find((i) => i.id === 'default'));
  await s1.api.tools.exportTool(a);
  const out = path.join(dir, 'books', 'budget-vba');
  fs.appendFileSync(path.join(out, 'Formatting.bas'), "' edited by hand\r\n");
  const old = new Date(Date.now() - 60000);
  for (const f of fs.readdirSync(out)) fs.utimesSync(path.join(out, f), old, old);
  const s2 = activate(dir);
  s2.ui.quick.push((items) => { assert.match(items[0].label, /Re-export to budget-vba/); return items[0]; });
  const r = await s2.api.tools.exportTool(a);
  assert.deepEqual(r.written, ['Formatting.bas']);
  assert.equal(r.same.length, 4);
  assert.ok(fs.statSync(path.join(out, 'LineItem.cls')).mtimeMs < Date.now() - 30000, 'unchanged file not rewritten');
  assert.doesNotMatch(fs.readFileSync(path.join(out, 'Formatting.bas'), 'utf8'), /edited by hand/);
  assert.match(s2.log.panel.posted.find((m) => m.type === 'export').summary, /^1 file written, 4 unchanged/);
});
test('Pro export: a file without VBA says so and uses no try; codepage encoding setting writes windows-1252', async () => {
  const dir = tmpWs();
  const s = activate(dir);
  await s.api.tools.exportTool(uriOf(s, dir, 'no-macros.xlsm'));
  assert.match(s.log.info.at(-1).text, /has no VBA modules, so there is nothing to export/);
  assert.equal(s.api.pro.left('exportModules'), 1);
  s.config.exportEncoding = 'codepage';
  s.ui.quick.push((items) => items[0]);
  await s.api.tools.exportTool(uriOf(s, dir, 'SimpleMacro.xls'));
  const m = JSON.parse(fs.readFileSync(path.join(dir, 'books', 'SimpleMacro-vba', '.vba-export.json'), 'utf8'));
  assert.equal(m.encoding, 'windows code page 1252');
  assert.equal(m.modules.length, 5);
});
test('Pro scan: file A shows the severity table and counts the try; reports save as Markdown and JSON; file B opens the Pro panel', async () => {
  const dir = tmpWs();
  const s = activate(dir);
  const rep = await s.api.tools.scanTool(uriOf(s, dir, 'budget.xlsm'));
  const posted = s.log.panel.posted.find((m) => m.type === 'scan');
  assert.equal(posted.report.level, 'high');
  assert.ok(posted.report.findings.some((f) => f.keyword === 'Workbook_Open'));
  assert.equal(rep.findings.length, posted.report.findings.length);
  assert.equal(s.api.pro.left('scan'), 0);
  await s.log.panel.send({ type: 'save', format: 'md' });
  await s.log.panel.send({ type: 'save', format: 'json' });
  const md = fs.readFileSync(path.join(dir, 'books', 'budget-macro-scan.md'), 'utf8');
  assert.match(md, /^# Macro security scan: budget\.xlsm/);
  assert.match(md, /\| high \| Auto-run \| Workbook_Open \|/);
  const js = JSON.parse(fs.readFileSync(path.join(dir, 'books', 'budget-macro-scan.json'), 'utf8'));
  assert.equal(js.file, 'budget.xlsm');
  assert.equal(js.level, 'high');
  assert.ok(s.log.panel.posted.some((m) => m.type === 'saved' && m.file === 'budget-macro-scan.json'));
  const n = s.log.posted.length;
  assert.equal(await s.api.tools.scanTool(uriOf(s, dir, 'SimpleMacro.xls')), undefined);
  assert.ok(proPanel(s, 'Macro security scan'));
  assert.equal(s.log.posted.length, n, 'no second report');
});
test('Pro scan: a file with nothing to scan uses no try; a broken file shows its error and uses no try', async () => {
  const dir = tmpWs();
  const s = activate(dir);
  await s.api.tools.scanTool(uriOf(s, dir, 'no-macros.xlsm'));
  assert.match(s.log.info.at(-1).text, /nothing to scan/);
  await s.api.tools.scanTool(uriOf(s, dir, 'not-office.xls'));
  assert.match(s.log.error.at(-1).text, /^not-office\.xls: This file is plain text/);
  assert.equal(s.api.pro.left('scan'), 1);
});
test('Pro compare: file A with B shows the module table, the diff opens; then file C opens the Pro panel', async () => {
  const dir = tmpWs();
  const s = activate(dir);
  s.ui.quick.push((items) => { assert.ok(!items.some((i) => i.label === 'budget.xlsm'), 'the file itself is not offered'); return items.find((i) => i.label === 'budget-v2.xlsm'); });
  const c = await s.api.tools.compareTool(uriOf(s, dir, 'budget.xlsm'));
  assert.equal(s.log.lastPick.opts.ignoreFocusOut, true);
  const posted = s.log.panel.posted.find((m) => m.type === 'compare');
  assert.deepEqual([posted.sum.changed, posted.sum.onlyA, posted.sum.onlyB, posted.sum.same], [1, 1, 1, 3]);
  assert.equal(posted.rows[0].name, 'BudgetReport');
  assert.equal(c.rows.length, 6);
  assert.equal(s.api.pro.left('compare'), 0);
  await s.log.panel.send({ type: 'diff', name: 'BudgetReport' });
  const diff = s.log.commands.find((x) => x[0] === 'vscode.diff');
  assert.equal(diff[1].scheme, 'vba-macro');
  assert.match(diff[1].path, /^\/budget\.xlsm\/BudgetReport\.bas$/);
  assert.match(diff[2].path, /^\/budget-v2\.xlsm\/BudgetReport\.bas$/);
  assert.match(diff[3], /^BudgetReport: budget\.xlsm ↔ budget-v2\.xlsm$/);
  const n = s.log.posted.length;
  assert.equal(await s.api.tools.compareTool(uriOf(s, dir, 'SimpleMacro.xls')), undefined);
  assert.ok(proPanel(s, 'Compare macros'));
  assert.equal(s.log.posted.length, n);
});
test('Pro tries survive a reload: a new extension host with the same state still has no try left for each used tool', async () => {
  const dir = tmpWs();
  const s = activate(dir);
  s.ui.quick.push((items) => items[0]);
  await s.api.tools.exportTool(uriOf(s, dir, 'budget.xlsm'));
  await s.api.tools.scanTool(uriOf(s, dir, 'budget.xlsm'));
  s.ui.quick.push((items) => items.find((i) => i.label === 'budget-v2.xlsm'));
  await s.api.tools.compareTool(uriOf(s, dir, 'budget.xlsm'));
  assert.deepEqual(s.api.pro.tries(), { exportModules: 0, scan: 0, compare: 0 });
  const r = activate(dir, s.state);
  assert.deepEqual(r.api.pro.tries(), { exportModules: 0, scan: 0, compare: 0 });
  await r.api.tools.scanTool(uriOf(r, dir, 'budget.xlsm'));
  assert.ok(proPanel(r, 'Macro security scan'));
  assert.match(proPanel(r, 'Macro security scan').opts.detail, /You used the free try of Macro security scan on this computer\./);
  assert.equal(r.log.panel, undefined);
});
test('Pro tools from the Explorer wrapper run on the clicked file (ctx surface); no file: a plain hint, no try', async () => {
  const dir = tmpWs();
  const s = activate(dir);
  await s.commands.get('vbaViewer.zovoPro.scan')(uriOf(s, dir, 'SimpleMacro.docm'));
  assert.ok(s.log.panel.posted.find((m) => m.type === 'scan' && m.file === 'SimpleMacro.docm'));
  const s2 = activate(dir);
  await s2.commands.get('vbaViewer.exportModules')();
  assert.match(s2.log.info.at(-1).text, /^Open the macros of an Office file first/);
  assert.equal(s2.api.pro.left('exportModules'), 1);
});

// ---------- README and text ----------
test('README: first screen has the main use, a screenshot of this product, the free list and the price; by Zovo; no rival names', () => {
  const md = fs.readFileSync(path.join(ROOT, 'README.md'), 'utf8');
  const first = md.split('\n').slice(0, 30).join('\n');
  assert.match(first, /^# VBA Viewer & Macro Extractor/);
  assert.match(first, /images\/screenshot\.png/);
  assert.match(first, /Zovo Pro: USD 19 once/);
  assert.match(md, /by Zovo/);
  assert.doesNotMatch(md, /serkonda7|spences10|XVBA|excel-live-server|aferri|excel-vba-sync|olevba|oletools/i);
  assert.ok(fs.existsSync(path.join(ROOT, 'images', 'screenshot.png')));
});
test('support links name this listing and related Zovo tools; walkthrough cross-sells listings that exist', () => {
  const support = require('../src/support');
  assert.equal(support.SELF_ID, 'zovo.zovo-vba-viewer');
  assert.ok(support.RELATED.length >= 3 && !support.RELATED.includes(support.SELF_ID));
  assert.equal(support.PRO.getPro, 'vbaViewer.getPro');
  const steps = JSON.stringify(pkg.contributes.walkthroughs);
  for (const id of ['zovo.table-viewer-vscode', 'zovo.docx-viewer-vscode', 'zovo.pptx-slide-viewer', 'zovo.zovo-hex-viewer']) assert.ok(steps.includes(`command:extension.open?%5B%22${id}%22%5D`), id);
  for (const m of pkg.contributes.walkthroughs[0].steps.map((x) => x.media.markdown).filter(Boolean)) assert.ok(fs.existsSync(path.join(ROOT, m)), m);
});
