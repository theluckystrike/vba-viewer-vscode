'use strict';
// Reads the VBA project of an Office file without Office and without running anything: finds the project storage
// (vbaProject.bin in OOXML packages, _VBA_PROJECT_CUR in .xls, Macros in .doc, the embedded storage in .ppt), parses
// the dir stream ([MS-OVBA] 2.3.4.2) and the PROJECT stream (2.3.1), and decompresses each module's source.
// Written for this extension from the public [MS-OVBA] specification.
const { readCfb, isCfb, CfbError } = require('./cfb');
const { decompress } = require('./ovba');
const { openZip, isZip } = require('./zip');
const { findPptVba } = require('./ppt');
const XLM = require('./xlm');

const CODEPAGES = {
  874: 'windows-874', 932: 'shift_jis', 936: 'gbk', 949: 'euc-kr', 950: 'big5', 1200: 'utf-16le', 1250: 'windows-1250', 1251: 'windows-1251',
  1252: 'windows-1252', 1253: 'windows-1253', 1254: 'windows-1254', 1255: 'windows-1255', 1256: 'windows-1256', 1257: 'windows-1257',
  1258: 'windows-1258', 10000: 'macintosh', 10007: 'x-mac-cyrillic', 20866: 'koi8-r', 21866: 'koi8-u', 28591: 'iso-8859-1', 28592: 'iso-8859-2',
  28595: 'iso-8859-5', 28597: 'iso-8859-7', 51932: 'euc-jp', 54936: 'gb18030', 65001: 'utf-8',
};
const decoders = new Map();
function codepageLabel(cp) { return CODEPAGES[cp] || 'windows-1252'; }
/** Text of MBCS bytes in a Windows code page (falls back to windows-1252 for pages the runtime does not know). */
function decodeCp(bytes, cp) {
  const label = codepageLabel(cp);
  let d = decoders.get(label);
  if (!d) {
    try { d = new TextDecoder(label); } catch (e) { d = new TextDecoder('windows-1252'); }
    decoders.set(label, d);
  }
  return d.decode(bytes);
}

const SYSKIND = { 0: '16-bit Windows', 1: '32-bit Windows', 2: 'Macintosh', 3: '64-bit Windows' };
const TYPE_LABEL = { document: 'Document module', standard: 'Standard module', class: 'Class module', form: 'Form module' };
const TYPE_EXT = { document: '.cls', standard: '.bas', class: '.cls', form: '.frm' };

/** Parses the decompressed dir stream. Returns { project, references, modules, warnings }. */
function parseDir(dir) {
  const project = { sysKind: null, lcid: null, lcidInvoke: null, codePage: 1252, name: '', docString: '', helpFile: '', helpContext: 0, libFlags: 0, versionMajor: 0, versionMinor: 0, constants: '', compatVersion: null };
  const references = [];
  const modules = [];
  const warnings = [];
  let p = 0;
  let cur = null;
  let ref = null;
  let moduleCount = null;
  const str = (data) => decodeCp(data, project.codePage);
  const u16 = (data) => data.toString('utf16le');
  while (p + 6 <= dir.length) {
    const id = dir.readUInt16LE(p);
    let size = dir.readUInt32LE(p + 2);
    p += 6;
    if (id === 0x0009) size = 6; // PROJECTVERSION: Reserved (4) is followed by 6 bytes, not by Reserved bytes
    if (p + size > dir.length) { warnings.push(`The dir stream is cut off in record 0x${id.toString(16).padStart(4, '0')}.`); break; }
    const data = dir.subarray(p, p + size);
    p += size;
    switch (id) {
      case 0x0001: project.sysKind = data.readUInt32LE(0); break;
      case 0x004a: project.compatVersion = data.readUInt32LE(0); break;
      case 0x0002: project.lcid = data.readUInt32LE(0); break;
      case 0x0014: project.lcidInvoke = data.readUInt32LE(0); break;
      case 0x0003: project.codePage = data.readUInt16LE(0); break;
      case 0x0004: project.name = str(data); break;
      case 0x0005: project.docString = str(data); break;
      case 0x0040: if (data.length) project.docString = u16(data); break;
      case 0x0006: project.helpFile = str(data); break;
      case 0x003d: break; // HelpFile2
      case 0x0007: project.helpContext = data.readUInt32LE(0); break;
      case 0x0008: project.libFlags = data.readUInt32LE(0); break;
      case 0x0009: project.versionMajor = data.readUInt32LE(0); project.versionMinor = data.readUInt16LE(4); break;
      case 0x000c: project.constants = str(data); break;
      case 0x003c: if (data.length) project.constants = u16(data); break;
      // references
      case 0x0016: ref = { name: str(data), kind: '', libid: '' }; references.push(ref); break;
      case 0x003e: if (ref && data.length) ref.name = u16(data); break;
      case 0x0033: if (!ref) { ref = { name: '', kind: '', libid: '' }; references.push(ref); } ref.kind = 'control'; ref.libid = str(data); break;
      case 0x002f: if (!ref) { ref = { name: '', kind: '', libid: '' }; references.push(ref); } ref.kind = 'control'; if (!ref.libid) ref.libid = libidOf(data, 0, str); break;
      case 0x0030: if (ref) ref.libid = libidOf(data, 0, str) || ref.libid; ref = null; break;
      case 0x000d: if (!ref) { ref = { name: '', kind: '', libid: '' }; references.push(ref); } ref.kind = 'registered'; ref.libid = libidOf(data, 0, str); ref = null; break;
      case 0x000e: if (!ref) { ref = { name: '', kind: '', libid: '' }; references.push(ref); } ref.kind = 'project'; ref.libid = libidOf(data, 0, str); ref = null; break;
      // modules
      case 0x000f: moduleCount = data.readUInt16LE(0); ref = null; break;
      case 0x0013: break; // PROJECTCOOKIE
      case 0x0019: cur = { name: str(data), nameUnicode: '', streamName: '', docString: '', offset: 0, helpContext: 0, cookie: 0, procedural: false, readOnly: false, private: false }; modules.push(cur); break;
      case 0x0047: if (cur && data.length) cur.nameUnicode = u16(data); break;
      case 0x001a: if (cur) cur.streamName = str(data); break;
      case 0x0032: if (cur && data.length) cur.streamName = u16(data); break;
      case 0x001c: if (cur) cur.docString = str(data); break;
      case 0x0048: if (cur && data.length) cur.docString = u16(data); break;
      case 0x0031: if (cur) cur.offset = data.readUInt32LE(0); break;
      case 0x001e: if (cur) cur.helpContext = data.readUInt32LE(0); break;
      case 0x002c: if (cur) cur.cookie = data.readUInt16LE(0); break;
      case 0x0021: if (cur) cur.procedural = true; break;
      case 0x0022: if (cur) cur.procedural = false; break;
      case 0x0025: if (cur) cur.readOnly = true; break;
      case 0x0028: if (cur) cur.private = true; break;
      case 0x002b: cur = null; break;
      case 0x0010: p = dir.length; break;
      default: break;
    }
  }
  for (const m of modules) { if (m.nameUnicode) m.name = m.nameUnicode; if (!m.streamName) m.streamName = m.name; delete m.nameUnicode; }
  if (moduleCount != null && moduleCount !== modules.length) warnings.push(`The dir stream says ${moduleCount} modules but lists ${modules.length}.`);
  return { project, references: references.filter((r) => r.name || r.libid), modules, warnings };
}

function libidOf(data, at, str) {
  if (data.length < at + 4) return '';
  const n = data.readUInt32LE(at);
  if (n > data.length - at - 4) return '';
  return str(data.subarray(at + 4, at + 4 + n));
}

/** Parses the PROJECT stream text: module kinds, protection, host extender info. */
function parseProjectStream(text) {
  const out = { kinds: new Map(), lines: [], protected: false, locked: false, id: '', name: '' };
  for (const raw of String(text).split(/\r\n|\n|\r/)) {
    const line = raw.trim();
    if (!line) continue;
    if (line.startsWith('[')) break;
    out.lines.push(line);
    const m = /^([A-Za-z]+)=(.*)$/.exec(line);
    if (!m) continue;
    const k = m[1].toLowerCase();
    const v = m[2];
    if (k === 'document') out.kinds.set(v.split('/')[0].toLowerCase(), 'document');
    else if (k === 'module') out.kinds.set(v.toLowerCase(), 'standard');
    else if (k === 'class') out.kinds.set(v.toLowerCase(), 'class');
    else if (k === 'baseclass') out.kinds.set(v.toLowerCase(), 'form');
    else if (k === 'id') out.id = v.replace(/"/g, '');
    else if (k === 'name') out.name = v.replace(/"/g, '');
    else if (k === 'dpb') {
      // DPB holds the password data, encrypted as in [MS-OVBA] 2.4.3. Without a password it is one 0x00 byte.
      const data = decryptData(v.replace(/"/g, ''));
      out.protected = !data || data.length !== 1 || data[0] !== 0;
    } else if (k === 'cmg' || k === 'gc') { /* protection state and visibility, encrypted */ }
  }
  return out;
}

/** Data Encryption ([MS-OVBA] 2.4.3.3): the plain bytes of a CMG, DPB or GC value, or null when it is malformed. */
function decryptData(hex) {
  if (!/^(?:[0-9A-Fa-f]{2})+$/.test(hex)) return null;
  const e = Buffer.from(hex, 'hex');
  if (e.length < 3) return null;
  const seed = e[0];
  const versionEnc = e[1];
  const projKeyEnc = e[2];
  let unenc1 = seed ^ projKeyEnc;
  let enc1 = projKeyEnc;
  let enc2 = versionEnc;
  let p = 3;
  const next = () => {
    if (p >= e.length) throw new Error('short');
    const be = e[p++];
    const bt = be ^ ((enc2 + unenc1) & 0xff);
    enc2 = enc1; enc1 = be; unenc1 = bt;
    return bt;
  };
  try {
    for (let i = 0; i < ((seed & 6) >> 1); i++) next();
    let len = 0;
    for (let i = 0; i < 4; i++) len |= next() << (8 * i);
    if (len < 0 || len > e.length) return null;
    const out = Buffer.alloc(len);
    for (let i = 0; i < len; i++) out[i] = next();
    return out;
  } catch (err) { return null; }
}

const ATTR_RE = /^\s*Attribute\s+[\w.]+\s*=/i;
const PROC_RE = /^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([\p{L}_][\p{L}\p{N}_]*)/iu;

/** Splits module source into attribute lines and code; lists procedures with their line in the code view. */
function analyzeSource(source) {
  const lines = String(source).split(/\r\n|\n|\r/);
  if (lines.length && lines[lines.length - 1] === '') lines.pop();
  const attributes = [];
  const code = [];
  const procedures = [];
  for (const l of lines) {
    if (ATTR_RE.test(l)) { attributes.push(l.trim()); continue; }
    code.push(l);
    const m = PROC_RE.exec(l);
    if (m) procedures.push({ kind: m[1].replace(/\s+/g, ' '), name: m[2], line: code.length });
  }
  const attr = {};
  for (const a of attributes) {
    const m = /^Attribute\s+(VB_\w+)\s*=\s*(.*)$/i.exec(a);
    if (m) attr[m[1]] = m[2].replace(/^"(.*)"$/, '$1');
  }
  return { code, attributes, attr, procedures };
}

function guessKind(m, attr) {
  if (m.procedural) return 'standard';
  const base = attr.VB_Base || '';
  const guids = (base.match(/\{[0-9A-Fa-f-]{36}\}/g) || []).length;
  if (guids >= 2) return 'form';
  if (guids === 1 && /^0\{/.test(base)) return 'document';
  return 'class';
}

/** Finds the storage that holds a VBA project (a VBA storage with a dir stream). Breadth-first, 4 levels deep. */
function findVbaStorage(cfb) {
  const queue = [{ e: cfb.root, depth: 0, path: [] }];
  while (queue.length) {
    const { e, depth, path } = queue.shift();
    const vba = cfb.child(e, 'VBA');
    if (vba && vba.type === 1 && cfb.child(vba, 'dir')) return { storage: e, vba, path };
    if (depth >= 4) continue;
    for (const c of cfb.children(e)) if (c.type === 1) queue.push({ e: c, depth: depth + 1, path: path.concat(c.name) });
  }
  return null;
}

/** Reads the VBA project in a compound file (a vbaProject.bin, .xls, .doc or embedded storage). */
function readProject(cfb) {
  const found = findVbaStorage(cfb);
  if (!found) return null;
  const warnings = [];
  const dirRaw = cfb.read(cfb.child(found.vba, 'dir'));
  let dir;
  try { dir = decompress(dirRaw); } catch (e) { throw new Error(`The VBA dir stream could not be decompressed: ${e.message}`); }
  const parsed = parseDir(dir);
  warnings.push(...parsed.warnings);
  const projEntry = cfb.child(found.storage, 'PROJECT');
  const projText = projEntry && projEntry.type === 2 ? decodeCp(cfb.read(projEntry), parsed.project.codePage) : '';
  const ps = parseProjectStream(projText);
  const modules = parsed.modules.map((m) => {
    const out = { name: m.name, streamName: m.streamName, docString: m.docString, readOnly: m.readOnly, private: m.private, offset: m.offset, type: 'class', source: '', size: 0, compressedSize: 0, error: '', attributes: [], attr: {}, procedures: [], lines: 0, frame: '' };
    const se = cfb.child(found.vba, m.streamName);
    if (!se || se.type !== 2) { out.error = `The module stream "${m.streamName}" is missing from the VBA storage.`; }
    else {
      try {
        const raw = cfb.read(se);
        out.compressedSize = raw.length;
        if (m.offset >= raw.length) throw new Error(`the source offset ${m.offset} is past the end of the ${raw.length}-byte stream`);
        const bytes = decompress(raw, m.offset);
        out.size = bytes.length;
        out.source = decodeCp(bytes, parsed.project.codePage);
      } catch (e) { out.error = `The source of this module could not be read: ${e.message}.`; }
    }
    const a = analyzeSource(out.source);
    out.attributes = a.attributes;
    out.attr = a.attr;
    out.procedures = a.procedures;
    out.lines = a.code.length;
    out.type = ps.kinds.get(m.name.toLowerCase()) || guessKind(m, a.attr);
    if (out.type === 'form') {
      const ds = cfb.child(found.storage, m.streamName) || cfb.child(found.storage, m.name);
      const fr = ds && ds.type === 1 ? cfb.child(ds, '\u0003VBFrame') : null;
      if (fr && fr.type === 2) { try { out.frame = decodeCp(cfb.read(fr), parsed.project.codePage); } catch (e) { /* no frame text */ } }
      if (ds && ds.type === 1) out.controls = formControls(cfb, ds);
    }
    return out;
  });
  const order = { document: 0, standard: 1, class: 2, form: 3 };
  modules.sort((x, y) => (order[x.type] - order[y.type]) || x.name.localeCompare(y.name, 'en', { sensitivity: 'base' }));
  return {
    storagePath: found.path.join('/'),
    project: { ...parsed.project, codePageLabel: codepageLabel(parsed.project.codePage), sysKindLabel: SYSKIND[parsed.project.sysKind] || '', id: ps.id },
    references: parsed.references,
    modules,
    protected: ps.protected,
    projectLines: ps.lines,
    warnings: warnings.concat(cfb.warnings || []),
  };
}

/** Child storages of a form's designer storage (frames and multipage pages); single controls live in its "o" stream, which is not parsed. */
function formControls(cfb, ds) {
  return cfb.children(ds).filter((c) => c.type === 1).map((c) => c.name).slice(0, 200);
}

function hostOf(names) {
  const s = names.map((n) => n.toLowerCase());
  if (s.some((n) => n.startsWith('xl/'))) return 'Excel';
  if (s.some((n) => n.startsWith('word/'))) return 'Word';
  if (s.some((n) => n.startsWith('ppt/'))) return 'PowerPoint';
  return '';
}

/**
 * Reads a file's macros. bytes: Buffer. name: file name (only for messages and the host guess).
 * Returns { ok: true, container, host, vbaPath, hasVba, project, references, modules, xlm, protected, warnings, note }
 * or { ok: false, error }.
 */
function extract(bytes, name = '') {
  const b = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const base = { ok: true, container: '', host: '', vbaPath: '', hasVba: false, project: null, references: [], modules: [], xlm: [], xlmNames: [], protected: false, warnings: [], note: '' };
  try {
    if (b.length === 0) return { ok: false, error: 'The file is empty (0 bytes).' };
    if (isZip(b)) {
      const zip = openZip(b);
      const out = { ...base, container: 'ooxml', host: hostOf(zip.names) };
      if (!zip.has('[Content_Types].xml')) return { ok: false, error: 'This is a ZIP file, but not an Office Open XML package (it has no [Content_Types].xml).' };
      const part = zip.names.find((n) => /(^|\/)vbaProject\.bin$/i.test(n));
      try { const x = XLM.fromOoxml(zip); out.xlm = x.sheets; out.xlmNames = x.names; } catch (e) { out.warnings.push(`Excel 4.0 macro sheets could not be listed: ${e.message}`); }
      if (!part) { out.note = 'This package has no vbaProject.bin part, so it holds no VBA macros.'; return out; }
      out.vbaPath = part;
      const cfb = readCfb(zip.read(part));
      const proj = readProject(cfb);
      if (!proj) { out.note = `The part ${part} has no VBA storage with a dir stream.`; return out; }
      return finish(out, proj);
    }
    if (isCfb(b)) {
      const cfb = readCfb(b);
      const out = { ...base, container: 'ole' };
      if (cfb.child(cfb.root, 'EncryptedPackage') || cfb.child(cfb.root, 'EncryptionInfo')) {
        return { ok: false, encrypted: true, error: 'This file is encrypted with a password to open. Its macros cannot be read until Office saves it without that password.' };
      }
      if (cfb.child(cfb.root, 'Workbook') || cfb.child(cfb.root, 'Book')) out.host = 'Excel';
      else if (cfb.child(cfb.root, 'WordDocument')) out.host = 'Word';
      else if (cfb.child(cfb.root, 'PowerPoint Document')) out.host = 'PowerPoint';
      if (out.host === 'Excel') {
        try { const x = XLM.fromBiff8(cfb); out.xlm = x.sheets; out.xlmNames = x.names; } catch (e) { out.warnings.push(`Excel 4.0 macro sheets could not be listed: ${e.message}`); }
      }
      let proj = readProject(cfb);
      if (proj) {
        out.vbaPath = proj.storagePath || '(root)';
        if (!proj.storagePath) out.container = 'vbaProject';
        return finish(out, proj);
      }
      if (out.host === 'PowerPoint') {
        const emb = findPptVba(cfb);
        if (emb) {
          proj = readProject(readCfb(emb.bytes));
          if (proj) { out.vbaPath = `PowerPoint Document, embedded storage at byte ${emb.offset}`; return finish(out, proj); }
        }
      }
      out.note = out.host ? `This ${out.host} file has no VBA project storage, so it holds no VBA macros.` : 'This compound file has no VBA project storage, so it holds no VBA macros.';
      return out;
    }
    const head = b.subarray(0, 512).toString('latin1');
    if (/^\s*<(?:!DOCTYPE\s+html|html|\?xml)/i.test(head)) return { ok: false, error: 'This file is HTML or XML text with an Office file name, not a binary Office file, so it has no VBA project.' };
    if (/^[\x09\x0a\x0d\x20-\x7e]*$/.test(head)) return { ok: false, error: 'This file is plain text (for example CSV) with an Office file name, not an Office file, so it has no VBA project.' };
    return { ok: false, error: 'This file is neither an OLE compound file nor a ZIP package, so it is not an Office file this viewer can read. It may be damaged or of another type.' };
  } catch (e) {
    const msg = String((e && e.message) || e);
    return { ok: false, error: e instanceof CfbError || /ZIP|compound|cut off|damaged/i.test(msg) ? msg : `The file could not be read: ${msg}` };
  }
}

function finish(out, proj) {
  out.hasVba = proj.modules.length > 0;
  out.project = proj.project;
  out.references = proj.references;
  out.modules = proj.modules;
  out.protected = proj.protected;
  out.warnings = out.warnings.concat(proj.warnings);
  if (!out.hasVba) out.note = 'The VBA project is there but has no modules.';
  return out;
}

/** Text of a module as the VBA editor exports it (.bas, .cls or .frm). */
function exportText(m) {
  let src = String(m.source || '').replace(/\r?\n/g, '\r\n').replace(/\r(?!\n)/g, '\r\n');
  // The VBA editor leaves these three module attributes out of exported class, document and form files.
  if (m.type !== 'standard') src = src.replace(/^Attribute VB_(?:Base|TemplateDerived|Customizable) = .*\r\n/gm, '');
  const body = src.endsWith('\r\n') || !src ? src : `${src}\r\n`;
  if (m.type === 'standard') return body;
  if (m.type === 'form') {
    const frame = String(m.frame || '').replace(/\r?\n/g, '\r\n').replace(/\s+$/, '');
    const head = frame && /^VERSION/i.test(frame) ? `${frame}\r\n` : `VERSION 5.00\r\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ${m.name}\r\nEnd\r\n`;
    return head + body;
  }
  return `VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n${body}`;
}

module.exports = { decryptData, extract, parseDir, parseProjectStream, analyzeSource, readProject, findVbaStorage, decodeCp, codepageLabel, exportText, TYPE_LABEL, TYPE_EXT };
