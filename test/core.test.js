'use strict';
// Core tests: MS-OVBA compression against the examples in section 3.2 of the specification, the compound file and
// ZIP readers, dir and PROJECT stream parsing, real Office files (Apache POI test files, VBA-JSON), Excel 4.0 macro
// sheets, PowerPoint embedded storage, the security scan, export files and the module compare.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const { decompress, compress, OvbaError, bitCount } = require('../src/core/ovba');
const { readCfb, isCfb, CfbError } = require('../src/core/cfb');
const { openZip } = require('../src/core/zip');
const V = require('../src/core/vba');
const XLM = require('../src/core/xlm');
const S = require('../src/core/scan');
const X = require('../src/core/export');
const DF = require('../src/core/diff');
const B = require('./support/build');

const FIX = path.join(__dirname, 'fixtures');
const REAL = path.join(FIX, 'real');
const read = (f) => fs.readFileSync(path.join(FIX, f));
const hex = (s) => Buffer.from(s.replace(/\s+/g, ''), 'hex');
const ex = (f) => V.extract(read(f), path.basename(f));
const mod = (r, n) => r.modules.find((m) => m.name === n);
const codeOf = (m) => V.analyzeSource(m.source).code.join('\n');

// ---------- MS-OVBA 3.2 examples ----------
const SPEC = [
  ['3.2.1 no compression', 'abcdefghijklmnopqrstuv.', '01 19 B0 00 61 62 63 64 65 66 67 68 00 69 6A 6B 6C 6D 6E 6F 70 00 71 72 73 74 75 76 2E'],
  ['3.2.2 normal compression', '#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa', '01 2F B0 00 23 61 61 61 62 63 64 65 82 66 00 70 61 67 68 69 6A 01 38 08 61 6B 6C 00 20 6D 6E 6F 70 06 71 02 70 04 00 72 73 74 75 76 10 77 78 79 7A 00 2C'],
  ['3.2.3 maximum compression', 'a'.repeat(73), '01 03 B0 02 61 45 00'],
];
for (const [name, plain, comp] of SPEC) {
  test(`MS-OVBA ${name}: decompression gives the example string`, () => {
    assert.equal(decompress(hex(comp)).toString('latin1'), plain);
  });
  test(`MS-OVBA ${name}: compression gives the example bytes`, () => {
    assert.deepEqual(compress(Buffer.from(plain, 'latin1')), hex(comp));
  });
}
test('MS-OVBA: copy token bit count follows CopyToken Help (minimum 4, up to 12)', () => {
  assert.equal(bitCount(1), 4);
  assert.equal(bitCount(16), 4);
  assert.equal(bitCount(17), 5);
  assert.equal(bitCount(2048), 11);
  assert.equal(bitCount(2049), 12);
  assert.equal(bitCount(4096), 12);
});
test('MS-OVBA: round trip over several chunks, incompressible data as raw chunks, and long runs', () => {
  const text = Buffer.from(Array.from({ length: 3000 }, (_, i) => `    Cells(${i % 37}, 2).Value = "row ${i}"\r\n`).join(''), 'latin1');
  assert.ok(text.length > 3 * 4096);
  const c = compress(text);
  assert.ok(c.length < text.length / 3);
  assert.deepEqual(decompress(c), text);
  const crypto = require('crypto');
  const noise = Buffer.concat(Array.from({ length: 313 }, (_, i) => crypto.createHash('sha256').update(`block ${i}`).digest())).subarray(0, 10000);
  const cn = compress(noise);
  assert.equal(cn.readUInt16LE(1) & 0x8000, 0, 'first chunk stored raw');
  assert.deepEqual(decompress(cn).subarray(0, noise.length), noise);
  const run = Buffer.alloc(9000, 0x41);
  assert.deepEqual(decompress(compress(run)), run);
});
test('MS-OVBA: bad signature byte, bad chunk signature and a copy token before the chunk start are errors', () => {
  assert.throws(() => decompress(hex('02 19 B0 00 61')), OvbaError);
  assert.throws(() => decompress(hex('01 19 A0 00 61 62')), /bad signature/);
  assert.throws(() => decompress(hex('01 03 B0 01 00 10')), /before the start of its chunk/);
  assert.throws(() => decompress(Buffer.alloc(0)), /empty/);
});

// ---------- compound file ----------
test('CFB: reads a written file with mini and regular streams, nested storages, case-insensitive names', () => {
  const big = Buffer.alloc(9000, 7);
  const small = Buffer.from('hello');
  const f = B.writeCfb({ children: [{ name: 'A', children: [{ name: 'Big', data: big }, { name: 'small', data: small }, { name: 'Zero', data: Buffer.alloc(0) }] }, { name: 'B', data: Buffer.from('b') }] });
  assert.ok(isCfb(f));
  const c = readCfb(f);
  assert.deepEqual(c.children(c.root).map((e) => e.name).sort(), ['A', 'B']);
  assert.deepEqual(c.read(c.find(['a', 'BIG'])), big);
  assert.deepEqual(c.read(c.find(['A', 'small'])), small);
  assert.equal(c.read(c.find(['A', 'Zero'])).length, 0);
  assert.equal(c.find(['A', 'none']), null);
});
test('CFB: not a compound file, cut off, and a looping sector chain give plain errors', () => {
  assert.throws(() => readCfb(Buffer.from('PK\u0003\u0004 not ole')), CfbError);
  const f = B.writeCfb({ children: [{ name: 'S', data: Buffer.alloc(5000, 1) }] });
  assert.throws(() => readCfb(f.subarray(0, 300)), /shorter than its 512-byte header/);
  const c = readCfb(f.subarray(0, 2048));
  assert.throws(() => c.read(c.find(['S'])), /cut off|past the end/);
  // make the stream's chain point to itself
  const loop = Buffer.from(f);
  const cc = readCfb(loop);
  const s = cc.find(['S']).start;
  loop.writeUInt32LE(s, 512 + s * 4);
  assert.throws(() => readCfb(loop).read(readCfb(loop).find(['S'])), /loops back/);
});

// ---------- zip ----------
test('ZIP: lists and inflates parts; a cut file has no central directory', () => {
  const z = openZip(read('../../samples/sample-macros.xlsm'));
  assert.ok(z.has('[Content_Types].xml'));
  assert.ok(z.has('XL/VBAPROJECT.BIN'));
  assert.ok(isCfb(z.read('xl/vbaProject.bin')));
  assert.throws(() => openZip(read('broken/cut-off.xlsm')), /no central directory/);
});

// ---------- dir and PROJECT streams ----------
test('dir stream: project, code page, references and module records (PROJECTVERSION has 6 data bytes)', () => {
  const c = readCfb(read('password.bin'));
  const dir = V.parseDir(decompress(c.read(c.find(['VBA', 'dir']))));
  assert.equal(dir.project.name, 'Locked');
  assert.equal(dir.project.codePage, 1252);
  assert.equal(dir.project.versionMajor, 1);
  assert.equal(dir.project.sysKind, 1);
  assert.equal(dir.references[0].name, 'stdole');
  assert.match(dir.references[0].libid, /^\*\\G\{00020430-/);
  assert.equal(dir.modules.length, 1);
  assert.equal(dir.modules[0].streamName, 'Module1');
  assert.equal(dir.modules[0].procedural, true);
  assert.deepEqual(dir.warnings, []);
});
test('PROJECT stream: module kinds; the DPB value is decrypted ([MS-OVBA] 2.4.3) to tell a password', () => {
  const p = V.parseProjectStream('ID="{X}"\r\nDocument=ThisWorkbook/&H00000000\r\nModule=Module1\r\nClass=Class1\r\nBaseClass=UserForm1\r\nName="VBAProject"\r\n\r\n[Workspace]\r\nModule=Ignored\r\n');
  assert.equal(p.kinds.get('thisworkbook'), 'document');
  assert.equal(p.kinds.get('module1'), 'standard');
  assert.equal(p.kinds.get('class1'), 'class');
  assert.equal(p.kinds.get('userform1'), 'form');
  assert.equal(p.kinds.size, 4);
  const id = '{5D9A0C21-6B1E-4C3F-9A77-3E0B2F1C8D44}';
  assert.deepEqual([...V.decryptData(B.encryptData(id, Buffer.from([0]), 0x2c))], [0]);
  assert.equal(V.decryptData(B.encryptData(id, Buffer.alloc(29, 9), 0x16)).length, 29);
  // a real DPB from Apache POI SimpleMacro.xls (no password)
  assert.deepEqual([...V.decryptData('A2A0881B581C581C58')], [0]);
  assert.equal(V.parseProjectStream('DPB="A2A0881B581C581C58"').protected, false);
});
test('module source: Attribute lines are split off; procedures are listed with their line in the code view', () => {
  const a = V.analyzeSource('Attribute VB_Name = "M"\r\nSub A()\r\nAttribute A.VB_Description = "x"\r\nEnd Sub\r\nPrivate Function Größe() As Long\r\nEnd Function\r\nPublic Property Get P()\r\nEnd Property\r\n');
  assert.deepEqual(a.attributes, ['Attribute VB_Name = "M"', 'Attribute A.VB_Description = "x"']);
  assert.equal(a.attr.VB_Name, 'M');
  assert.deepEqual(a.procedures.map((p) => [p.kind, p.name, p.line]), [['Sub', 'A', 1], ['Function', 'Größe', 3], ['Property Get', 'P', 5]]);
});

// ---------- real files (Apache POI test files, Apache-2.0; VBA-JSON, MIT) ----------
test('real .xls (POI SimpleMacro.xls): _VBA_PROJECT_CUR, 5 modules, Module1 equals the published source', () => {
  const r = V.extract(fs.readFileSync(path.join(REAL, 'SimpleMacro.xls')));
  assert.equal(r.ok, true);
  assert.equal(r.container, 'ole');
  assert.equal(r.host, 'Excel');
  assert.equal(r.vbaPath, '_VBA_PROJECT_CUR');
  assert.deepEqual(r.modules.map((m) => `${m.type}:${m.name}`), ['document:Sheet1', 'document:Sheet2', 'document:Sheet3', 'document:ThisWorkbook', 'standard:Module1']);
  const want = fs.readFileSync(path.join(REAL, 'SimpleMacro-xls.vba'), 'latin1').replace(/\r\n/g, '\n').trimEnd();
  assert.equal(codeOf(mod(r, 'Module1')).trimEnd(), want);
  assert.equal(r.protected, false);
  assert.deepEqual(r.references.map((x) => x.name), ['stdole', 'Office']);
});
test('real .xlsm, .doc, .docm, .pptm (POI): the VBA project in each container', () => {
  const x = V.extract(fs.readFileSync(path.join(REAL, 'SimpleMacro.xlsm')));
  assert.equal(x.vbaPath, 'xl/vbaProject.bin');
  assert.match(codeOf(mod(x, 'Module1')), /ActiveCell\.FormulaR1C1 = "This is a macro workbook"/);
  const d = V.extract(fs.readFileSync(path.join(REAL, 'SimpleMacro.doc')));
  assert.equal(d.host, 'Word');
  assert.equal(d.vbaPath, 'Macros');
  assert.deepEqual(d.modules.map((m) => m.name), ['ThisDocument', 'Module1']);
  const dm = V.extract(fs.readFileSync(path.join(REAL, 'SimpleMacro.docm')));
  assert.equal(dm.vbaPath, 'word/vbaProject.bin');
  assert.equal(codeOf(mod(dm, 'Module1')), codeOf(mod(d, 'Module1')));
  const pm = V.extract(fs.readFileSync(path.join(REAL, 'SimpleMacro.pptm')));
  assert.equal(pm.host, 'PowerPoint');
  assert.match(codeOf(mod(pm, 'Module1')), /This is a macro slideshow/);
});
test('real .ppt (POI SimpleMacro.ppt): the VBA storage embedded in the PowerPoint Document stream is found and read', () => {
  const r = V.extract(fs.readFileSync(path.join(REAL, 'SimpleMacro.ppt')));
  assert.equal(r.ok, true);
  assert.equal(r.host, 'PowerPoint');
  assert.match(r.vbaPath, /^PowerPoint Document, embedded storage/);
  const pptx = V.extract(fs.readFileSync(path.join(REAL, 'SimpleMacro.pptm')));
  assert.equal(codeOf(mod(r, 'Module1')), codeOf(mod(pptx, 'Module1')));
});
test('real workbook (VBA-JSON specs, MIT): 10 modules; exported JsonConverter.bas equals the published file (the repo stores it with LF line ends)', () => {
  const r = V.extract(fs.readFileSync(path.join(REAL, 'VBA-JSON-Specs.xlsm')));
  assert.equal(r.modules.length, 10);
  assert.equal(mod(r, 'Dictionary').type, 'class');
  assert.equal(mod(r, 'SpecRunner').type, 'document');
  const want = fs.readFileSync(path.join(REAL, 'JsonConverter.bas'));
  const got = X.exportFiles(r, 'codepage').files.find((f) => f.file === 'JsonConverter.bas');
  assert.ok(got);
  assert.equal(got.bytes.toString('latin1').replace(/\r\n/g, '\n'), want.toString('latin1'));
  assert.ok(got.bytes.includes('\r\n'), 'the export keeps CRLF as the VBA editor writes it');
  assert.equal(mod(r, 'JsonConverter').procedures.length, 24);
});
test('real files with no VBA: .xlsb (POI Simple.xlsb) and an Excel 4.0 sheet without VBA (POI xlmmacro.xlsm)', () => {
  const b = V.extract(fs.readFileSync(path.join(REAL, 'Simple.xlsb')));
  assert.equal(b.ok, true);
  assert.equal(b.hasVba, false);
  assert.match(b.note, /no vbaProject\.bin part/);
  const x = V.extract(fs.readFileSync(path.join(REAL, 'xlmmacro.xlsm')));
  assert.equal(x.hasVba, false);
  assert.deepEqual(x.xlm.map((s) => [s.name, s.state, s.kind]), [['Macro1', 'visible', 'Excel 4.0 macro sheet']]);
});

// ---------- made-up fixtures ----------
test('sample workbook: 5 modules of 3 kinds in the VBA editor order, sources and attributes', () => {
  const r = V.extract(read('../../samples/sample-macros.xlsm'));
  assert.equal(r.project.name, 'BudgetMacros');
  assert.deepEqual(r.modules.map((m) => `${m.type}:${m.name}`), ['document:Sheet1', 'document:ThisWorkbook', 'standard:BudgetReport', 'standard:Formatting', 'class:LineItem']);
  const br = mod(r, 'BudgetReport');
  assert.deepEqual(br.procedures.map((p) => p.name), ['BuildSummary', 'MonthTotal', 'VatAmount']);
  assert.equal(br.attributes[0], 'Attribute VB_Name = "BudgetReport"');
  assert.equal(mod(r, 'LineItem').attr.VB_PredeclaredId, 'False');
});
test('forms: BaseClass module type, the VBFrame text, and the .frm export', () => {
  const r = ex('forms.xlsm');
  const f = mod(r, 'EntryForm');
  assert.equal(f.type, 'form');
  assert.match(f.frame, /Caption\s+=\s+"New budget line"/);
  const t = V.exportText(f);
  assert.match(t, /^VERSION 5\.00\r\nBegin \{C62A69F0-16DC-11CE-9E98-00AA00574A4F\} EntryForm ?\r\n/);
  assert.match(t, /End\r\nAttribute VB_Name = "EntryForm"\r\nAttribute VB_GlobalNameSpace = False\r\n/);
  assert.doesNotMatch(t, /VB_Base|VB_TemplateDerived|VB_Customizable/);
});
test('password: the project is marked, and the source is still read', () => {
  const r = ex('password.bin');
  assert.equal(r.container, 'vbaProject');
  assert.equal(r.protected, true);
  assert.match(mod(r, 'Module1').source, /still readable/);
});
test('code page 1251: module name, project name and source in Cyrillic', () => {
  const r = ex('cp1251.bin');
  assert.equal(r.project.codePage, 1251);
  assert.equal(r.project.name, 'Проект');
  const m = mod(r, 'Отчет');
  assert.ok(m);
  assert.match(m.source, /Сумма по месяцам/);
  assert.deepEqual(m.procedures.map((p) => p.name), ['Итог']);
  const files = X.exportFiles(r, 'codepage');
  assert.equal(files.encodingUsed, 'windows code page 1251');
  assert.match(require('../src/core/vba').decodeCp(files.files[0].bytes, 1251), /MsgBox "Готово"/);
});
test('Word-style compound file: Macros storage, host Word', () => {
  const r = ex('word-style.doc');
  assert.equal(r.host, 'Word');
  assert.equal(r.vbaPath, 'Macros');
  assert.equal(mod(r, 'NewMacros').lines, 3);
});
test('synthetic .ppt: an uncompressed ExOleObjStg inside a container record is found too', () => {
  const inner = B.buildVbaProject({ modules: [{ name: 'Slides', type: 'standard', source: 'Attribute VB_Name = "Slides"\r\nSub Go()\r\nEnd Sub\r\n' }] });
  const atom = (ver, inst, type, body) => { const h = Buffer.alloc(8); h.writeUInt16LE(ver | (inst << 4), 0); h.writeUInt16LE(type, 2); h.writeUInt32LE(body.length, 4); return Buffer.concat([h, body]); };
  const z = zlib.deflateSync(inner);
  const sz = Buffer.alloc(4); sz.writeUInt32LE(inner.length);
  const stream = Buffer.concat([atom(0xf, 0, 0x03e8, atom(0, 0, 0x03e9, Buffer.alloc(40))), atom(0, 1, 0x1011, Buffer.concat([sz, z]))]);
  const ppt = B.writeCfb({ children: [{ name: 'PowerPoint Document', data: Buffer.concat([stream, Buffer.alloc(5000)]) }, { name: 'Current User', data: Buffer.alloc(20) }] });
  const r = V.extract(ppt);
  assert.equal(r.host, 'PowerPoint');
  assert.equal(mod(r, 'Slides').lines, 2);
  const plain = B.writeCfb({ children: [{ name: 'PowerPoint Document', data: Buffer.concat([atom(0, 0, 0x1011, inner), Buffer.alloc(100)]) }] });
  assert.equal(V.extract(plain).modules.length, 1);
});
test('Excel 4.0: very hidden macro sheet with formulas and an Auto_Open name (OOXML); BIFF8 BoundSheet records; BIFF12 BrtBundleSh', () => {
  const r = ex('xlm-hidden.xlsm');
  assert.deepEqual(r.xlm.map((s) => [s.name, s.state, s.formulas.length]), [['Macro1', 'very hidden', 2]]);
  assert.deepEqual(r.xlmNames, [{ name: 'Auto_Open', ref: 'Macro1!$A$1' }]);
  // BIFF8 globals: BOF, BoundSheet8 (worksheet), BoundSheet8 (macro sheet, hidden), Lbl builtin Auto_Open, EOF
  const recB = (type, body) => { const h = Buffer.alloc(4); h.writeUInt16LE(type, 0); h.writeUInt16LE(body.length, 2); return Buffer.concat([h, body]); };
  const bs = (name, state, dt) => Buffer.concat([Buffer.from([0, 0, 0, 0, state, dt, name.length, 0]), Buffer.from(name, 'latin1')]);
  const lbl = Buffer.concat([Buffer.from([0x20, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]), Buffer.from([1])]);
  const wb = Buffer.concat([recB(0x0809, Buffer.from([0, 6, 5, 0])), recB(0x0085, bs('Data', 0, 0)), recB(0x0085, bs('Macro1', 1, 1)), recB(0x0018, lbl), recB(0x000a, Buffer.alloc(0)), recB(0x0085, bs('Late', 0, 1))]);
  const xls = B.writeCfb({ children: [{ name: 'Workbook', data: Buffer.concat([wb, Buffer.alloc(4096)]) }] });
  const rx = V.extract(xls);
  assert.equal(rx.host, 'Excel');
  assert.deepEqual(rx.xlm.map((s) => [s.name, s.state]), [['Macro1', 'hidden']]);
  assert.deepEqual(rx.xlmNames.map((n) => n.name), ['Auto_Open']);
  // BIFF12: BrtBundleSh (156) with a 2-byte record type
  const ws = (s) => { const b = Buffer.alloc(4); b.writeUInt32LE(s.length); return Buffer.concat([b, Buffer.from(s, 'utf16le')]); };
  const body = Buffer.concat([Buffer.from([2, 0, 0, 0, 1, 0, 0, 0]), ws('rId2'), ws('Macro2')]);
  const recb = Buffer.concat([Buffer.from([0x9c, 0x01, body.length]), body]);
  assert.deepEqual(XLM.biff12Sheets(recb), [{ state: 2, relId: 'rId2', name: 'Macro2' }]);
});
test('no VBA, encrypted and broken files: clear empty state or a plain error, never a throw', () => {
  const n = ex('no-macros.xlsm');
  assert.equal(n.ok, true);
  assert.equal(n.hasVba, false);
  assert.match(n.note, /no vbaProject\.bin part/);
  assert.match(ex('broken/cut-off.xlsm').error, /no central directory/);
  assert.match(ex('broken/not-office.xls').error, /plain text \(for example CSV\)/);
  assert.match(ex('broken/random.doc').error, /neither an OLE compound file nor a ZIP package/);
  assert.match(ex('broken/cut-vbaProject.bin').error, /cut off|damaged/);
  assert.match(V.extract(Buffer.alloc(0)).error, /empty/);
  const enc = B.writeCfb({ children: [{ name: 'EncryptionInfo', data: Buffer.alloc(200) }, { name: 'EncryptedPackage', data: Buffer.alloc(5000) }] });
  assert.match(V.extract(enc).error, /encrypted with a password to open/);
  assert.match(V.extract(Buffer.from('<html><body>report</body></html>')).error, /HTML or XML text/);
  // a module whose stream is damaged: the others still show
  const bin = B.buildVbaProject({ modules: [{ name: 'Good', type: 'standard', source: 'Sub A()\r\nEnd Sub\r\n' }, { name: 'Bad', type: 'standard', source: 'Sub B()\r\nEnd Sub\r\n' }] });
  const c = readCfb(bin);
  const e = c.find(['VBA', 'Bad']);
  const z = Buffer.from(bin);
  const mini = readCfb(z);
  const rootStart = mini.root.start;
  const pos = 512 + rootStart * 512 + e.start * 64;
  z[pos] = 0x07; // signature byte no longer 0x01
  const r = V.extract(z);
  assert.equal(mod(r, 'Good').error, '');
  assert.match(mod(r, 'Bad').error, /could not be read: The compressed data does not start with the signature byte/);
});

// ---------- scan ----------
test('scan: sample workbook finds the auto-run entry points, CreateObject, the URL, Environ and the hex string', () => {
  const rep = S.scan(V.extract(read('../../samples/sample-macros.xlsm')), { fileName: 'sample-macros.xlsm' });
  const k = (t, kw) => rep.findings.find((f) => f.type === t && f.keyword === kw);
  assert.equal(k('Auto-run', 'Workbook_Open').severity, 'high');
  assert.deepEqual(k('Auto-run', 'Workbook_Open').places, [{ module: 'ThisWorkbook', line: 3 }]);
  assert.ok(k('Auto-run', 'Workbook_BeforeClose'));
  assert.equal(k('Suspicious', 'CreateObject').severity, 'medium');
  assert.equal(k('IOC', 'https://example.com/budget/help').description, 'URL');
  assert.equal(k('Suspicious', 'Environ').severity, 'low');
  const h = rep.findings.find((f) => f.type === 'Hex string');
  assert.equal(h.decoded, 'Hello, world!');
  assert.equal(rep.level, 'high');
  assert.equal(rep.autoRun, true);
  assert.match(rep.summary, /runs by itself when the file opens or closes\. No calls that start programs/);
  assert.equal(rep.modules, 5);
});
test('scan: risky calls, comments ignored, Base64 and Chr() obfuscation, IP addresses, file names, Excel 4.0', () => {
  // inert text written for this test; it is never run
  const src = ['Attribute VB_Name = "T"', 'Sub AutoOpen()', '    \' Shell "never.exe" in a comment is ignored', '    x = Shell("notepad.exe", 1)', '    Set o = CreateObject("WScript.Shell")',
    '    u = "http://192.0.2.10/files/setup.exe"', '    s = "cG93ZXJzaGVsbCAtbm9wIC1jIGVjaG8gaGk="', '    y = ' + Array.from({ length: 21 }, (_, i) => `Chr(${65 + (i % 26)})`).join(' & '),
    '    Declare PtrSafe Function VirtualAlloc Lib "kernel32" () As LongPtr', 'End Sub'].join('\r\n');
  const r = { modules: [{ name: 'T', source: src, error: '' }], xlm: [{ name: 'M1', state: 'very hidden', kind: 'Excel 4.0 macro sheet', formulas: [{ cell: 'A1', formula: '=EXEC("calc")' }] }], xlmNames: [{ name: 'Auto_Open', ref: '' }], protected: true };
  const rep = S.scan(r);
  const kws = rep.findings.map((f) => `${f.severity}:${f.type}:${f.keyword}`);
  for (const w of ['high:Auto-run:AutoOpen', 'high:Suspicious:Shell', 'high:Suspicious:WScript.Shell', 'medium:Suspicious:CreateObject', 'medium:IOC:192.0.2.10', 'medium:IOC:setup.exe', 'medium:IOC:notepad.exe', 'high:Base64 string:cG93ZXJzaGVsbCAtbm9wIC1jIGVjaG8gaGk=', 'medium:Obfuscation:Chr()', 'high:Suspicious:Memory API', 'high:Excel 4.0:EXEC', 'high:Excel 4.0:M1', 'high:Auto-run:Auto_Open', 'info:Info:VBA password']) assert.ok(kws.includes(w), `${w} in ${kws.join(', ')}`);
  assert.equal(rep.findings.find((f) => f.keyword === 'Shell').count, 1, 'the comment does not count');
  assert.equal(rep.findings.find((f) => f.type === 'Base64 string').decoded, 'powershell -nop -c echo hi');
  assert.ok(!kws.includes('medium:IOC:never.exe'));
  assert.match(rep.summary, /runs by itself.*start programs/);
  const md = S.toMarkdown(rep, { file: 't.xlsm', date: '2026-11-15 20:00' });
  assert.match(md, /^# Macro security scan: t\.xlsm\n/);
  assert.match(md, /\| high \| Suspicious \| Shell \| Starts another program \| 1 \| T:3 \|/);
  assert.match(md, /nothing was run/);
});
test('scan: a file with only harmless code has no findings; string splitting keeps doubled quotes', () => {
  const rep = S.scan({ modules: [{ name: 'M', source: 'Sub Add()\r\n    x = 1 + 2 \' Shell in a comment\r\nEnd Sub\r\n' }], xlm: [] });
  assert.equal(rep.level, 'none');
  assert.match(rep.summary, /No auto-run entry points and no risky calls/);
  assert.deepEqual(S.splitLine('a = "say ""hi""" \' note').strings, ['say "hi"']);
  assert.equal(S.splitLine('Rem all of this').code, '');
});

// ---------- export ----------
test('export: .bas, .cls and .frm names and texts; a re-export plan rewrites only changed files and names orphans', () => {
  const r = ex('forms.xlsm');
  const { files } = X.exportFiles(r);
  assert.deepEqual(files.map((f) => f.file), ['Sheet1.cls', 'ThisWorkbook.cls', 'BudgetReport.bas', 'Formatting.bas', 'LineItem.cls', 'EntryForm.frm']);
  const cls = files.find((f) => f.file === 'LineItem.cls').bytes.toString('utf8');
  assert.match(cls, /^VERSION 1\.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = "LineItem"\r\nAttribute VB_GlobalNameSpace = False\r\n/);
  assert.match(files.find((f) => f.file === 'BudgetReport.bas').bytes.toString('utf8'), /^Attribute VB_Name = "BudgetReport"\r\nOption Explicit\r\n/);
  const first = X.planExport(files, new Map(), null);
  assert.equal(first.write.length, 6);
  const disk = new Map(files.map((f) => [f.file.toLowerCase(), f.bytes]));
  disk.set('oldmodule.bas', Buffer.from('x'));
  const changed = files.map((f) => (f.file === 'Formatting.bas' ? { ...f, bytes: Buffer.concat([f.bytes, Buffer.from("' new\r\n")]) } : f));
  const again = X.planExport(changed, disk, { modules: [...files.map((f) => ({ file: f.file })), { file: 'OldModule.bas' }] });
  assert.deepEqual(again.write.map((f) => f.file), ['Formatting.bas']);
  assert.equal(again.same.length, 5);
  assert.deepEqual(again.orphans, ['OldModule.bas']);
  const m = JSON.parse(X.manifest('forms.xlsm', files, 'utf8', '2026-11-15T00:00:00Z'));
  assert.equal(m.modules.length, 6);
  assert.match(m.modules[0].sha256, /^[0-9a-f]{64}$/);
  assert.equal(X.encodeCp('Größe €', 1252).toString('hex'), '4772f6df652080');
  assert.equal(X.encodeCp('x', 932), null);
});

// ---------- compare ----------
test('compare: changed, only in one file and same modules, with line counts', () => {
  const a = V.extract(read('../../samples/sample-macros.xlsm'));
  const b = ex('sample-macros-v2.xlsm');
  const c = DF.compareResults(a, b);
  const st = Object.fromEntries(c.rows.map((r) => [r.name, r]));
  assert.equal(st.BudgetReport.status, 'changed');
  assert.equal(st.BudgetReport.removed, 1);
  assert.equal(st.BudgetReport.added, 2);
  assert.equal(st.Formatting.status, 'onlyA');
  assert.equal(st.Archive.status, 'onlyB');
  assert.equal(st.LineItem.status, 'same');
  assert.deepEqual([c.sum.changed, c.sum.onlyA, c.sum.onlyB, c.sum.same], [1, 1, 1, 3]);
  assert.equal(c.rows[0].name, 'BudgetReport');
  assert.deepEqual(DF.lineDelta(['a', 'b', 'c'], ['a', 'x', 'c', 'd']), { removed: 1, added: 2, exact: true });
});
