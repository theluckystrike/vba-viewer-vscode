'use strict';
// Reader for OLE compound files (Compound File Binary format, the container of .xls, .doc, .ppt and vbaProject.bin),
// written for this extension from the public [MS-CFB] specification. Read-only: it never writes to the file. Every
// sector chain is followed with a loop guard, so a damaged or hostile file ends in a plain error, never a hang.

const SIG = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const MAXREGSECT = 0xfffffffa;
const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;
const NOSTREAM = 0xffffffff;

class CfbError extends Error {}

function isCfb(b) {
  if (!b || b.length < 8) return false;
  for (let i = 0; i < 8; i++) if (b[i] !== SIG[i]) return false;
  return true;
}

/** Parses a compound file. b: Buffer or Uint8Array. Returns { root, entries, children(e), find(path), read(e), warnings }. */
function readCfb(input) {
  const b = Buffer.isBuffer(input) ? input : Buffer.from(input.buffer, input.byteOffset, input.byteLength);
  if (!isCfb(b)) throw new CfbError('This is not an OLE compound file (the first 8 bytes are not D0 CF 11 E0 A1 B1 1A E1).');
  if (b.length < 512) throw new CfbError('The compound file is cut off: it is shorter than its 512-byte header.');
  const warnings = [];
  const major = b.readUInt16LE(0x1a);
  const shift = b.readUInt16LE(0x1e);
  if (shift !== 9 && shift !== 12) throw new CfbError(`Unsupported sector size 2^${shift} in the compound file header.`);
  const ss = 1 << shift;
  const miniShift = b.readUInt16LE(0x20);
  const ms = 1 << (miniShift > 0 && miniShift < shift ? miniShift : 6);
  const nFat = b.readUInt32LE(0x2c);
  const firstDir = b.readUInt32LE(0x30);
  const cutoff = b.readUInt32LE(0x38) || 4096;
  const firstMiniFat = b.readUInt32LE(0x3c);
  const nMiniFat = b.readUInt32LE(0x40);
  const firstDifat = b.readUInt32LE(0x44);
  const nDifat = b.readUInt32LE(0x48);
  const nSectors = Math.max(0, Math.floor((b.length - ss) / ss) + ((b.length - ss) % ss ? 1 : 0));
  const per = ss / 4;
  const sectorOk = (s) => s < MAXREGSECT && s < nSectors;
  const sectorBuf = (s) => {
    const off = (s + 1) * ss;
    if (off >= b.length) return null;
    return b.subarray(off, Math.min(off + ss, b.length));
  };

  // DIFAT: the first 109 FAT sector numbers are in the header, the rest in a chain of DIFAT sectors.
  const fatSectors = [];
  for (let i = 0; i < 109 && fatSectors.length < nFat; i++) {
    const v = b.readUInt32LE(0x4c + i * 4);
    if (v < MAXREGSECT) fatSectors.push(v);
  }
  let d = firstDifat;
  const seenDifat = new Set();
  while (d < MAXREGSECT && fatSectors.length < nFat && seenDifat.size <= nDifat + 1) {
    if (seenDifat.has(d)) { warnings.push('The DIFAT chain loops; the rest of it was skipped.'); break; }
    seenDifat.add(d);
    const sb = sectorBuf(d);
    if (!sb || sb.length < ss) { warnings.push('The DIFAT chain points past the end of the file.'); break; }
    for (let j = 0; j < per - 1 && fatSectors.length < nFat; j++) {
      const v = sb.readUInt32LE(j * 4);
      if (v < MAXREGSECT) fatSectors.push(v);
    }
    d = sb.readUInt32LE((per - 1) * 4);
  }

  const fat = new Uint32Array(fatSectors.length * per).fill(FREESECT);
  fatSectors.forEach((s, i) => {
    const sb = sectorBuf(s);
    if (!sb) { warnings.push(`FAT sector ${s} is past the end of the file.`); return; }
    for (let j = 0; j < per && j * 4 + 4 <= sb.length; j++) fat[i * per + j] = sb.readUInt32LE(j * 4);
  });

  function chain(start, table, limit) {
    const out = [];
    const seen = new Set();
    let s = start;
    while (s < MAXREGSECT) {
      if (seen.has(s)) throw new CfbError('A sector chain in the compound file loops back on itself; the file is damaged.');
      if (s >= table.length) throw new CfbError('A sector chain in the compound file points outside its allocation table; the file is damaged or cut off.');
      seen.add(s);
      out.push(s);
      if (limit && out.length >= limit) break;
      s = table[s];
    }
    return out;
  }

  function readChain(start, size) {
    const want = size == null ? Infinity : size;
    const sectors = chain(start, fat, size == null ? 0 : Math.ceil(size / ss));
    const parts = [];
    let got = 0;
    for (const s of sectors) {
      if (got >= want) break;
      const sb = sectorBuf(s);
      if (!sb) throw new CfbError('A stream in the compound file reaches past the end of the file; the file is cut off.');
      parts.push(sb);
      got += sb.length;
    }
    const all = Buffer.concat(parts);
    if (size != null && all.length < size) throw new CfbError('A stream in the compound file is shorter than its directory entry says; the file is cut off.');
    return size == null ? all : all.subarray(0, size);
  }

  if (!sectorOk(firstDir)) throw new CfbError('The compound file directory starts outside the file; the file is damaged or cut off.');
  const dirBuf = readChain(firstDir, null);
  const entries = [];
  for (let off = 0; off + 128 <= dirBuf.length; off += 128) {
    const nameLen = dirBuf.readUInt16LE(off + 0x40);
    const name = nameLen >= 2 && nameLen <= 64 ? dirBuf.toString('utf16le', off, off + nameLen - 2) : '';
    const type = dirBuf[off + 0x42];
    const lo = dirBuf.readUInt32LE(off + 0x78);
    const hi = dirBuf.readUInt32LE(off + 0x7c);
    entries.push({
      id: entries.length, name, type,
      left: dirBuf.readUInt32LE(off + 0x44), right: dirBuf.readUInt32LE(off + 0x48), child: dirBuf.readUInt32LE(off + 0x4c),
      clsid: dirBuf.subarray(off + 0x50, off + 0x60).toString('hex'),
      start: dirBuf.readUInt32LE(off + 0x74),
      size: major === 3 ? lo : hi * 4294967296 + lo,
    });
  }
  const root = entries[0];
  if (!root || root.type !== 5) throw new CfbError('The compound file has no root entry; the file is damaged.');

  let miniFat = null;
  let miniStream = null;
  function loadMini() {
    if (miniFat) return;
    const mf = firstMiniFat < MAXREGSECT ? readChain(firstMiniFat, null) : Buffer.alloc(0);
    miniFat = new Uint32Array(Math.floor(mf.length / 4));
    for (let i = 0; i < miniFat.length; i++) miniFat[i] = mf.readUInt32LE(i * 4);
    miniStream = root.start < MAXREGSECT ? readChain(root.start, root.size) : Buffer.alloc(0);
    if (nMiniFat && miniFat.length === 0) warnings.push('The mini FAT is empty.');
  }

  function read(e) {
    if (!e || e.type !== 2) throw new CfbError('Not a stream.');
    if (e.size === 0) return Buffer.alloc(0);
    if (e.size > b.length * 2 + 65536) throw new CfbError(`The stream "${e.name}" claims ${e.size} bytes, more than the file can hold; the file is damaged.`);
    if (e.size >= cutoff) return readChain(e.start, e.size);
    loadMini();
    const sectors = chain(e.start, miniFat, Math.ceil(e.size / ms));
    const parts = [];
    for (const s of sectors) {
      const off = s * ms;
      if (off + ms > miniStream.length && off >= miniStream.length) throw new CfbError(`The stream "${e.name}" points outside the mini stream; the file is damaged.`);
      parts.push(miniStream.subarray(off, Math.min(off + ms, miniStream.length)));
    }
    const all = Buffer.concat(parts);
    if (all.length < e.size) throw new CfbError(`The stream "${e.name}" is cut off.`);
    return all.subarray(0, e.size);
  }

  const kids = new Map();
  /** Child entries of a storage (or the root), in directory tree order. */
  function children(e) {
    if (!e) return [];
    if (kids.has(e.id)) return kids.get(e.id);
    const out = [];
    const seen = new Set();
    const stack = [];
    let n = e.child;
    // In-order walk of the red-black tree of siblings, without recursion.
    while ((n !== NOSTREAM && n < entries.length) || stack.length) {
      while (n !== NOSTREAM && n < entries.length) {
        if (seen.has(n)) { warnings.push('The directory tree of the compound file loops; some entries were skipped.'); n = NOSTREAM; break; }
        seen.add(n);
        stack.push(n);
        n = entries[n].left;
      }
      if (!stack.length) break;
      const cur = stack.pop();
      if (entries[cur].type === 1 || entries[cur].type === 2) out.push(entries[cur]);
      n = entries[cur].right;
    }
    kids.set(e.id, out);
    return out;
  }
  /** Child by name (case-insensitive, as in the format). */
  function child(e, name) {
    const l = String(name).toLowerCase();
    return children(e).find((c) => c.name.toLowerCase() === l) || null;
  }
  /** Entry at a path of names from the root, e.g. ['Macros', 'VBA', 'dir']. */
  function find(path) {
    let e = root;
    for (const p of path) { e = child(e, p); if (!e) return null; }
    return e;
  }
  return { root, entries, children, child, find, read, warnings, major, sectorSize: ss };
}

module.exports = { readCfb, isCfb, CfbError, ENDOFCHAIN, FREESECT, NOSTREAM };
