'use strict';
// Export all modules: the file set (.bas, .cls, .frm) the VBA editor would write, as bytes, plus the plan of a
// re-export that only rewrites files whose content changed. Pure functions; src/tools.js does the disk work.
const crypto = require('crypto');
const { exportText, TYPE_EXT, decodeCp } = require('./vba');

const MANIFEST = '.vba-export.json';
const SINGLE_BYTE = new Set([874, 1250, 1251, 1252, 1253, 1254, 1255, 1256, 1257, 1258, 10000, 28591, 28592, 28595, 28597, 20866, 21866]);
const encoders = new Map();

/** Encodes text in a single-byte Windows code page (characters it cannot hold become "?"). null for other pages. */
function encodeCp(text, cp) {
  if (!SINGLE_BYTE.has(cp)) return null;
  let map = encoders.get(cp);
  if (!map) {
    map = new Map();
    const all = Buffer.alloc(256);
    for (let i = 0; i < 256; i++) all[i] = i;
    const s = decodeCp(all, cp);
    if (s.length === 256) for (let i = 255; i >= 0; i--) map.set(s[i], i);
    encoders.set(cp, map);
  }
  const out = Buffer.alloc(text.length * 2);
  let n = 0;
  for (const ch of text) { const v = map.get(ch); out[n++] = v == null ? 0x3f : v; }
  return out.subarray(0, n);
}

const safeName = (n) => String(n).replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').replace(/^\.+/, '_') || 'Module';
const sha = (b) => crypto.createHash('sha256').update(b).digest('hex');

/**
 * files for a result: [{ module, type, file, bytes, sha256 }]. encoding: 'utf8' (default) or 'codepage'.
 * Returns { files, encodingUsed, note }.
 */
function exportFiles(result, encoding = 'utf8') {
  const cp = result.project ? result.project.codePage : 1252;
  let used = 'utf8';
  let note = '';
  const taken = new Set();
  const files = [];
  for (const m of result.modules || []) {
    if (m.error && !m.source) continue;
    const base = safeName(m.name);
    const ext = TYPE_EXT[m.type] || '.cls';
    let file = base + ext;
    for (let i = 2; taken.has(file.toLowerCase()); i++) file = `${base}-${i}${ext}`;
    taken.add(file.toLowerCase());
    const text = exportText(m);
    let bytes = null;
    if (encoding === 'codepage') {
      bytes = encodeCp(text, cp);
      if (bytes) used = `windows code page ${cp}`;
      else if (!note) note = `Code page ${cp} is not a single-byte page, so the files are UTF-8.`;
    }
    if (!bytes) bytes = Buffer.from(text, 'utf8');
    files.push({ module: m.name, type: m.type, file, bytes, sha256: sha(bytes) });
  }
  return { files, encodingUsed: used, note };
}

/**
 * Plan against what is on disk. existing: Map(lowercase file name -> Buffer). oldManifest: parsed manifest or null.
 * Returns { write: [file], same: [file], orphans: [file names listed before whose module is gone] }.
 */
function planExport(files, existing, oldManifest) {
  const write = [];
  const same = [];
  for (const f of files) {
    const cur = existing.get(f.file.toLowerCase());
    if (cur && Buffer.compare(cur, f.bytes) === 0) same.push(f); else write.push({ ...f, isNew: !cur });
  }
  const now = new Set(files.map((f) => f.file.toLowerCase()));
  const orphans = oldManifest && Array.isArray(oldManifest.modules) ? oldManifest.modules.map((m) => m && m.file).filter((f) => typeof f === 'string' && !now.has(f.toLowerCase()) && existing.has(f.toLowerCase())) : [];
  return { write, same, orphans };
}

function manifest(sourceName, files, encodingUsed, date) {
  return `${JSON.stringify({ source: sourceName, exported: date, tool: 'VBA Viewer & Macro Extractor (Zovo)', encoding: encodingUsed, modules: files.map((f) => ({ name: f.module, type: f.type, file: f.file, sha256: f.sha256 })) }, null, 2)}\n`;
}

module.exports = { exportFiles, planExport, manifest, encodeCp, MANIFEST, safeName };
