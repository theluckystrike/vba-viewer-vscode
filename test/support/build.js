'use strict';
// Builders for test fixtures and the shipped samples: a compound file writer, a VBA project writer (dir stream,
// PROJECT stream, compressed module streams, as [MS-OVBA] says a writer must: no p-code, _VBA_PROJECT version 0xFFFF,
// cookies 0xFFFF), a ZIP writer and a minimal macro-enabled OOXML package. Not shipped in the extension.
const zlib = require('zlib');
const { compress } = require('../../src/core/ovba');

// ---------- compound file writer (version 3, 512-byte sectors) ----------
const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;
const FATSECT = 0xfffffffd;
const NOSTREAM = 0xffffffff;

function cmpName(a, b) {
  if (a.length !== b.length) return a.length - b.length;
  const A = a.toUpperCase();
  const B = b.toUpperCase();
  return A < B ? -1 : A > B ? 1 : 0;
}

/** tree: { children: [{ name, data: Buffer } | { name, children: [...] }] } -> Buffer */
function writeCfb(tree) {
  const entries = [{ name: 'Root Entry', type: 5, node: tree }];
  const queue = [0];
  while (queue.length) {
    const i = queue.shift();
    const node = entries[i].node;
    if (!node.children) continue;
    const kids = node.children.slice().sort((a, b) => cmpName(a.name, b.name));
    const ids = kids.map((k) => { entries.push({ name: k.name, type: k.children ? 1 : 2, node: k }); return entries.length - 1; });
    // balanced tree: complete levels black, the last partial level red
    const depthOf = new Map();
    const build = (lo, hi, depth) => {
      if (lo > hi) return NOSTREAM;
      const mid = (lo + hi) >> 1;
      const id = ids[mid];
      depthOf.set(id, depth);
      entries[id].left = build(lo, mid - 1, depth + 1);
      entries[id].right = build(mid + 1, hi, depth + 1);
      return id;
    };
    entries[i].child = build(0, ids.length - 1, 0);
    const maxDepth = Math.max(-1, ...depthOf.values());
    const full = Math.floor(Math.log2(ids.length + 1)); // number of complete levels
    for (const [id, d] of depthOf) entries[id].color = d >= full && maxDepth >= full ? 0 : 1;
    for (const id of ids) if (entries[id].type === 1) queue.push(id);
  }
  const mini = [];
  const miniFat = [];
  const big = [];
  for (const e of entries) {
    if (e.type !== 2) continue;
    const data = e.node.data || Buffer.alloc(0);
    e.size = data.length;
    if (data.length === 0) { e.start = ENDOFCHAIN; continue; }
    if (data.length < 4096) {
      const first = miniFat.length;
      const n = Math.ceil(data.length / 64);
      for (let k = 0; k < n; k++) miniFat.push(k === n - 1 ? ENDOFCHAIN : first + k + 1);
      const padded = Buffer.alloc(n * 64);
      data.copy(padded);
      mini.push(padded);
      e.start = first;
    } else big.push(e);
  }
  const miniBuf = Buffer.concat(mini);
  const dirSectors = Math.ceil((entries.length * 128) / 512);
  const miniFatSectors = Math.ceil((miniFat.length * 4) / 512);
  const miniStreamSectors = Math.ceil(miniBuf.length / 512);
  const bigSectors = big.reduce((s, e) => s + Math.ceil(e.size / 512), 0);
  const n = dirSectors + miniFatSectors + miniStreamSectors + bigSectors;
  let f = 1;
  while (f * 128 < n + f) f++;
  if (f > 109) throw new Error('fixture too large for the simple writer');
  const total = f + n;
  const fat = new Array(f * 128).fill(FREESECT);
  for (let i = 0; i < f; i++) fat[i] = FATSECT;
  let next = f;
  const alloc = (count) => { if (!count) return ENDOFCHAIN; const s = next; for (let k = 0; k < count; k++) fat[s + k] = k === count - 1 ? ENDOFCHAIN : s + k + 1; next += count; return s; };
  const firstDir = alloc(dirSectors);
  const firstMiniFat = alloc(miniFatSectors);
  const miniStart = alloc(miniStreamSectors);
  for (const e of big) e.start = alloc(Math.ceil(e.size / 512));
  const out = Buffer.alloc(512 + total * 512);
  Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(out, 0);
  out.writeUInt16LE(0x3e, 0x18); out.writeUInt16LE(3, 0x1a); out.writeUInt16LE(0xfffe, 0x1c); out.writeUInt16LE(9, 0x1e); out.writeUInt16LE(6, 0x20);
  out.writeUInt32LE(0, 0x28); out.writeUInt32LE(f, 0x2c); out.writeUInt32LE(firstDir, 0x30); out.writeUInt32LE(0, 0x34); out.writeUInt32LE(4096, 0x38);
  out.writeUInt32LE(miniFatSectors ? firstMiniFat : ENDOFCHAIN, 0x3c); out.writeUInt32LE(miniFatSectors, 0x40); out.writeUInt32LE(ENDOFCHAIN, 0x44); out.writeUInt32LE(0, 0x48);
  for (let i = 0; i < 109; i++) out.writeUInt32LE(i < f ? i : FREESECT, 0x4c + i * 4);
  const sec = (s) => 512 + s * 512;
  for (let i = 0; i < fat.length; i++) out.writeUInt32LE(fat[i] >>> 0, sec(Math.floor(i / 128)) + (i % 128) * 4);
  // directory
  const root = entries[0];
  root.start = miniStreamSectors ? miniStart : ENDOFCHAIN;
  root.size = miniBuf.length;
  for (let i = 0; i < dirSectors * 4; i++) {
    const off = sec(firstDir + Math.floor(i / 4)) + (i % 4) * 128;
    const e = entries[i];
    if (!e) { out.writeUInt32LE(NOSTREAM, off + 0x44); out.writeUInt32LE(NOSTREAM, off + 0x48); out.writeUInt32LE(NOSTREAM, off + 0x4c); continue; }
    const nm = Buffer.from(`${e.name}\u0000`, 'utf16le');
    nm.copy(out, off, 0, Math.min(64, nm.length));
    out.writeUInt16LE(Math.min(64, nm.length), off + 0x40);
    out[off + 0x42] = e.type;
    out[off + 0x43] = e.color == null ? 1 : e.color;
    out.writeUInt32LE(e.left == null ? NOSTREAM : e.left, off + 0x44);
    out.writeUInt32LE(e.right == null ? NOSTREAM : e.right, off + 0x48);
    out.writeUInt32LE(e.child == null ? NOSTREAM : e.child, off + 0x4c);
    out.writeUInt32LE(e.type === 1 || e.type === 5 ? 0 : (e.start >>> 0), off + 0x74);
    if (e.type === 5) out.writeUInt32LE(root.start >>> 0, off + 0x74);
    out.writeUInt32LE(e.type === 1 ? 0 : (e.size || 0), off + 0x78);
  }
  for (let i = 0; i < miniFat.length; i++) out.writeUInt32LE(miniFat[i] >>> 0, sec(firstMiniFat + Math.floor(i / 128)) + (i % 128) * 4);
  for (let i = miniFat.length; i < miniFatSectors * 128; i++) out.writeUInt32LE(FREESECT, sec(firstMiniFat + Math.floor(i / 128)) + (i % 128) * 4);
  if (miniStreamSectors) miniBuf.copy(out, sec(miniStart));
  for (const e of big) e.node.data.copy(out, sec(e.start));
  return out;
}

// ---------- VBA project writer ----------
function rec(id, data) { const h = Buffer.alloc(6); h.writeUInt16LE(id, 0); h.writeUInt32LE(data.length, 2); return Buffer.concat([h, data]); }
const u16 = (v) => { const b = Buffer.alloc(2); b.writeUInt16LE(v); return b; };
const u32 = (v) => { const b = Buffer.alloc(4); b.writeUInt32LE(v >>> 0); return b; };
const ansi = (s) => Buffer.from(s, 'latin1');
const uni = (s) => Buffer.from(s, 'utf16le');

function encryptData(projectId, data, seed = 0x4a) {
  let projKey = 0;
  for (const c of Buffer.from(projectId, 'latin1')) projKey = (projKey + c) & 0xff;
  const version = 2;
  const out = [seed, seed ^ version, seed ^ projKey];
  let unenc1 = projKey;
  let enc1 = seed ^ projKey;
  let enc2 = seed ^ version;
  const put = (t) => { const be = t ^ ((enc2 + unenc1) & 0xff); enc2 = enc1; enc1 = be; unenc1 = t; out.push(be); };
  for (let i = 0; i < ((seed & 6) >> 1); i++) put(7);
  for (let i = 0; i < 4; i++) put((data.length >> (8 * i)) & 0xff);
  for (const b of data) put(b);
  return Buffer.from(out).toString('hex').toUpperCase();
}

/**
 * o: { name, codePage, modules: [{ name, type: 'document'|'standard'|'class'|'form', source, frame? }], password?: bool, encode?(text)->Buffer }
 * Returns the vbaProject.bin bytes.
 */
function buildVbaProject(o) {
  const name = o.name || 'VBAProject';
  const cp = o.codePage || 1252;
  const enc = o.encode || ((s) => Buffer.from(s, 'latin1'));
  const parts = [];
  parts.push(rec(0x0001, u32(1)), rec(0x004a, u32(3)), rec(0x0002, u32(0x409)), rec(0x0014, u32(0x409)), rec(0x0003, u16(cp)), rec(0x0004, enc(name)));
  parts.push(rec(0x0005, Buffer.alloc(0)), rec(0x0040, Buffer.alloc(0)), rec(0x0006, Buffer.alloc(0)), rec(0x003d, Buffer.alloc(0)), rec(0x0007, u32(0)), rec(0x0008, u32(0)));
  const ver = Buffer.alloc(10); ver.writeUInt16LE(0x0009, 0); ver.writeUInt32LE(4, 2); ver.writeUInt32LE(1, 6); parts.push(ver, u16(0));
  parts.push(rec(0x000c, Buffer.alloc(0)), rec(0x003c, Buffer.alloc(0)));
  const libid = '*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation';
  parts.push(rec(0x0016, ansi('stdole')), rec(0x003e, uni('stdole')), rec(0x000d, Buffer.concat([u32(libid.length), ansi(libid), u32(0), u16(0)])));
  parts.push(rec(0x000f, u16(o.modules.length)), rec(0x0013, u16(0xffff)));
  const vba = [];
  for (const m of o.modules) {
    parts.push(rec(0x0019, enc(m.name)), rec(0x0047, uni(m.name)), rec(0x001a, enc(m.name)), rec(0x0032, uni(m.name)), rec(0x001c, Buffer.alloc(0)), rec(0x0048, Buffer.alloc(0)));
    parts.push(rec(0x0031, u32(0)), rec(0x001e, u32(0)), rec(0x002c, u16(0xffff)), rec(m.type === 'standard' ? 0x0021 : 0x0022, Buffer.alloc(0)), rec(0x002b, Buffer.alloc(0)));
    vba.push({ name: m.name, data: compress(enc(m.source)) });
  }
  parts.push(rec(0x0010, Buffer.alloc(0)));
  const dir = compress(Buffer.concat(parts));
  vba.push({ name: 'dir', data: dir }, { name: '_VBA_PROJECT', data: Buffer.from([0xcc, 0x61, 0xff, 0xff, 0x00, 0x00, 0x00]) });
  const id = '{5D9A0C21-6B1E-4C3F-9A77-3E0B2F1C8D44}';
  const lines = [`ID="${id}"`];
  for (const m of o.modules) {
    if (m.type === 'document') lines.push(`Document=${m.name}/&H00000000`);
    else if (m.type === 'standard') lines.push(`Module=${m.name}`);
    else if (m.type === 'class') lines.push(`Class=${m.name}`);
    else if (m.type === 'form') lines.push(`BaseClass=${m.name}`);
  }
  lines.push(`Name="${name}"`, 'HelpContextID="0"', 'VersionCompatible32="393222000"');
  lines.push(`CMG="${encryptData(id, Buffer.from([0, 0, 0, 0]), 0x26)}"`);
  lines.push(`DPB="${encryptData(id, o.password ? Buffer.alloc(29, 0x5a) : Buffer.from([0]), 0x2c)}"`);
  lines.push(`GC="${encryptData(id, Buffer.from([0xff]), 0x31)}"`);
  lines.push('', '[Host Extender Info]', '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000', '', '[Workspace]');
  for (const m of o.modules) lines.push(`${m.name}=0, 0, 0, 0, C`);
  const children = [{ name: 'VBA', children: vba }, { name: 'PROJECT', data: enc(`${lines.join('\r\n')}\r\n`) }];
  for (const m of o.modules) if (m.type === 'form') children.push({ name: m.name, children: [{ name: '\u0003VBFrame', data: enc(m.frame || `VERSION 5.00\r\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ${m.name}\r\nEnd\r\n`) }, { name: 'f', data: Buffer.alloc(0) }, { name: 'o', data: Buffer.alloc(0) }] });
  return writeCfb({ children: o.wrap ? [{ name: o.wrap, children }] : children });
}

// ---------- ZIP writer ----------
const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(b) { let c = 0xffffffff; for (const x of b) c = CRC[(c ^ x) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }
/** files: [[name, Buffer|string]] -> Buffer */
function writeZip(files) {
  const locals = [];
  const central = [];
  let off = 0;
  for (const [name, content] of files) {
    const data = Buffer.isBuffer(content) ? content : Buffer.from(content, 'utf8');
    const def = zlib.deflateRawSync(data);
    const nm = Buffer.from(name, 'utf8');
    const crc = crc32(data);
    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(0x800, 6); lh.writeUInt16LE(8, 8); lh.writeUInt16LE(0, 10); lh.writeUInt16LE(0x5a21, 12);
    lh.writeUInt32LE(crc, 14); lh.writeUInt32LE(def.length, 18); lh.writeUInt32LE(data.length, 22); lh.writeUInt16LE(nm.length, 26); lh.writeUInt16LE(0, 28);
    locals.push(lh, nm, def);
    const ch = Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6); ch.writeUInt16LE(0x800, 8); ch.writeUInt16LE(8, 10); ch.writeUInt16LE(0, 12); ch.writeUInt16LE(0x5a21, 14);
    ch.writeUInt32LE(crc, 16); ch.writeUInt32LE(def.length, 20); ch.writeUInt32LE(data.length, 24); ch.writeUInt16LE(nm.length, 28); ch.writeUInt32LE(off, 42);
    central.push(ch, nm);
    off += 30 + nm.length + def.length;
  }
  const cd = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(files.length, 8); end.writeUInt16LE(files.length, 10); end.writeUInt32LE(cd.length, 12); end.writeUInt32LE(off, 16);
  return Buffer.concat([...locals, cd, end]);
}

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
/**
 * Minimal macro-enabled Excel package. o: { vba?: Buffer, sheets?: [names], macroSheets?: [{ name, state, formulas: [[cell, formula]] }], autoOpen?: string }
 */
function buildXlsm(o = {}) {
  const sheets = o.sheets || ['Sheet1'];
  const macro = o.macroSheets || [];
  const ct = [`${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`,
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    o.vba ? '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>' : '',
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.${o.vba || macro.length ? 'macroEnabled.main+xml' : 'main+xml'}"/>`.replace('sheet.main+xml', 'sheet.main+xml'),
    ...sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`),
    ...macro.map((s, i) => `<Override PartName="/xl/macrosheets/sheet${i + 1}.xml" ContentType="application/vnd.ms-excel.macrosheet+xml"/>`),
    '</Types>'].join('');
  const fixCt = o.vba || macro.length ? ct : ct.replace('application/vnd.ms-excel.sheet.main+xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml');
  const rels = `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const wbRels = [`${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`,
    ...sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`),
    ...macro.map((s, i) => `<Relationship Id="rId${100 + i}" Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" Target="macrosheets/sheet${i + 1}.xml"/>`),
    o.vba ? '<Relationship Id="rId900" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>' : '',
    '</Relationships>'].join('');
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const wb = [`${XML}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>`,
    ...sheets.map((s, i) => `<sheet name="${esc(s)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`),
    ...macro.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${50 + i}"${s.state && s.state !== 'visible' ? ` state="${s.state}"` : ''} r:id="rId${100 + i}"/>`),
    '</sheets>', o.autoOpen ? `<definedNames><definedName name="_xlnm.Auto_Open">${esc(o.autoOpen)}</definedName></definedNames>` : '', '</workbook>'].join('');
  const sheet = (rows) => `${XML}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows || ''}</sheetData></worksheet>`;
  const files = [['[Content_Types].xml', fixCt], ['_rels/.rels', rels], ['xl/workbook.xml', wb], ['xl/_rels/workbook.xml.rels', wbRels]];
  sheets.forEach((s, i) => files.push([`xl/worksheets/sheet${i + 1}.xml`, sheet(i === 0 && o.cells ? o.cells : '')]));
  macro.forEach((s, i) => files.push([`xl/macrosheets/sheet${i + 1}.xml`, `${XML}<xm:macrosheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><sheetData>${(s.formulas || []).map(([cell, f]) => `<row><c r="${cell}"><f>${esc(f)}</f></c></row>`).join('')}</sheetData></xm:macrosheet>`]));
  if (o.vba) files.push(['xl/vbaProject.bin', o.vba]);
  return writeZip(files);
}

module.exports = { writeCfb, buildVbaProject, writeZip, buildXlsm, encryptData, crc32 };
