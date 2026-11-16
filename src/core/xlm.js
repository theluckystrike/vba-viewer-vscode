'use strict';
// Excel 4.0 (XLM) macro sheets: lists them with their visibility in .xlsm/.xltm/.xlam (workbook.xml), .xlsb
// (workbook.bin, BIFF12) and .xls (Workbook stream, BIFF8 and BIFF5), plus Auto_Open style names. For OOXML macro
// sheets the cell formulas are read too (they are plain XML text). Nothing is evaluated.

const MACRO_REL = /\/(xlMacrosheet|xlIntlMacrosheet)$/i;
const AUTO_RE = /^(?:_xlnm\.)?auto_(open|close|activate|deactivate)\b/i;
const STATE = { 0: 'visible', 1: 'hidden', 2: 'very hidden' };

const unxml = (s) => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (m, d) => String.fromCodePoint(+d)).replace(/&#x([0-9a-f]+);/gi, (m, h) => String.fromCodePoint(parseInt(h, 16))).replace(/&amp;/g, '&');
const attrOf = (tag, name) => { const m = new RegExp(`(?:^|\\s)${name}="([^"]*)"`).exec(tag); return m ? unxml(m[1]) : ''; };

function relsOf(zip, path) {
  const dir = path.replace(/[^/]*$/, '');
  const relPath = `${dir}_rels/${path.slice(dir.length)}.rels`;
  const buf = zip.read(relPath);
  const map = new Map();
  if (!buf) return map;
  const xml = buf.toString('utf8');
  for (const m of xml.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = attrOf(m[0], 'Id');
    let target = attrOf(m[0], 'Target');
    const type = attrOf(m[0], 'Type');
    if (!target.startsWith('/')) target = normalize(dir + target); else target = target.slice(1);
    map.set(id, { target, type });
  }
  return map;
}
function normalize(p) {
  const out = [];
  for (const s of p.split('/')) { if (s === '..') out.pop(); else if (s && s !== '.') out.push(s); }
  return out.join('/');
}

function formulasOf(xml, max = 5000) {
  const out = [];
  for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    if (!m[2]) continue;
    const f = /<f\b[^>]*>([\s\S]*?)<\/f>/.exec(m[2]);
    if (!f) continue;
    out.push({ cell: attrOf(m[1], 'r'), formula: unxml(f[1]) });
    if (out.length >= max) break;
  }
  return out;
}

/** OOXML: { sheets: [{ name, state, kind, part, formulas }], names: [{ name, ref }] } */
function fromOoxml(zip) {
  const sheets = [];
  const names = [];
  const wbXml = zip.read('xl/workbook.xml');
  if (wbXml) {
    const rels = relsOf(zip, 'xl/workbook.xml');
    const xml = wbXml.toString('utf8');
    for (const m of xml.matchAll(/<(?:\w+:)?sheet\b[^>]*>/g)) {
      const rid = attrOf(m[0], 'r:id') || (/\bid="([^"]*)"/.exec(m[0]) || [])[1] || '';
      const rel = rels.get(rid);
      if (!rel || !MACRO_REL.test(rel.type)) continue;
      const st = attrOf(m[0], 'state');
      const part = zip.read(rel.target);
      sheets.push({ name: attrOf(m[0], 'name'), state: st === 'veryHidden' ? 'very hidden' : st === 'hidden' ? 'hidden' : 'visible', kind: /Intl/i.test(rel.type) ? 'International macro sheet' : 'Excel 4.0 macro sheet', part: rel.target, formulas: part ? formulasOf(part.toString('utf8')) : [] });
    }
    for (const m of xml.matchAll(/<(?:\w+:)?definedName\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?definedName>/g)) {
      const n = attrOf(m[1], 'name');
      if (AUTO_RE.test(n)) names.push({ name: n.replace(/^_xlnm\./, ''), ref: unxml(m[2]) });
    }
    return { sheets, names };
  }
  const wbBin = zip.read('xl/workbook.bin');
  if (wbBin) {
    const rels = relsOf(zip, 'xl/workbook.bin');
    for (const sh of biff12Sheets(wbBin)) {
      const rel = rels.get(sh.relId);
      if (!rel || !MACRO_REL.test(rel.type)) continue;
      sheets.push({ name: sh.name, state: STATE[sh.state] || 'visible', kind: /Intl/i.test(rel.type) ? 'International macro sheet' : 'Excel 4.0 macro sheet', part: rel.target, formulas: [] });
    }
  }
  return { sheets, names };
}

/** BrtBundleSh records of a workbook.bin: [{ state, relId, name }] */
function biff12Sheets(b) {
  const out = [];
  let p = 0;
  while (p < b.length) {
    let type = b[p++];
    if (type & 0x80) { type = (type & 0x7f) | ((b[p++] & 0x7f) << 7); }
    let size = 0;
    for (let i = 0; i < 4; i++) { const v = b[p++]; size |= (v & 0x7f) << (7 * i); if (!(v & 0x80)) break; }
    if (p + size > b.length) break;
    if (type === 156 && size >= 12) {
      const state = b.readUInt32LE(p);
      let q = p + 8;
      const n1 = b.readUInt32LE(q); q += 4;
      let relId = '';
      if (n1 !== 0xffffffff) { relId = b.toString('utf16le', q, q + n1 * 2); q += n1 * 2; }
      const n2 = b.readUInt32LE(q); q += 4;
      out.push({ state, relId, name: b.toString('utf16le', q, q + n2 * 2) });
    }
    p += size;
  }
  return out;
}

/** .xls: BoundSheet8 records with dt = 1 (macro sheet) in the Workbook (BIFF8) or Book (BIFF5) stream. */
function fromBiff8(cfb) {
  const e = cfb.child(cfb.root, 'Workbook') || cfb.child(cfb.root, 'Book');
  const sheets = [];
  const names = [];
  if (!e || e.type !== 2) return { sheets, names };
  const s = cfb.read(e);
  let p = 0;
  let biff8 = true;
  let first = true;
  let encrypted = false;
  while (p + 4 <= s.length) {
    const type = s.readUInt16LE(p);
    const len = s.readUInt16LE(p + 2);
    const d = p + 4;
    if (d + len > s.length) break;
    if (first) { first = false; if (type === 0x0809 && len >= 2) biff8 = s.readUInt16LE(d) === 0x0600; }
    else if (type === 0x000a) break; // EOF of the globals substream
    if (type === 0x002f) encrypted = true;
    if (type === 0x0085 && len >= 8 && !encrypted) {
      const state = s[d + 4] & 0x03;
      const dt = s[d + 5];
      const cch = s[d + 6];
      let name = '';
      if (biff8) { const hi = s[d + 7] & 1; name = hi ? s.toString('utf16le', d + 8, d + 8 + cch * 2) : s.toString('latin1', d + 8, d + 8 + cch); }
      else name = s.toString('latin1', d + 7, d + 7 + cch);
      if (dt === 0x01) sheets.push({ name, state: STATE[state] || 'visible', kind: 'Excel 4.0 macro sheet', part: '', formulas: [] });
    }
    if (type === 0x0018 && len >= 15 && biff8 && !encrypted) {
      const flags = s.readUInt16LE(d);
      const cch = s[d + 3];
      const hi = s[d + 14] & 1;
      const raw = hi ? s.toString('utf16le', d + 15, d + 15 + cch * 2) : s.toString('latin1', d + 15, d + 15 + cch);
      let n = raw;
      if (flags & 0x20) n = ({ '\u0001': 'Auto_Open', '\u0002': 'Auto_Close', '\u000a': 'Auto_Activate', '\u000b': 'Auto_Deactivate' })[raw[0]] || '';
      if (n && AUTO_RE.test(n)) names.push({ name: n, ref: '' });
    }
    p = d + len;
  }
  return { sheets, names, encrypted };
}

module.exports = { fromOoxml, fromBiff8, biff12Sheets, formulasOf };
