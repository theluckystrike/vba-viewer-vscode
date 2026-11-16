'use strict';
// Small ZIP reader for Office Open XML packages (.xlsm, .docm, .pptm, ...), written for this extension. It reads the
// central directory and inflates only the entries asked for, with Node's zlib. Read-only.
const zlib = require('zlib');

class ZipError extends Error {}

function isZip(b) { return !!b && b.length >= 4 && b[0] === 0x50 && b[1] === 0x4b && (b[2] === 3 || b[2] === 5 || b[2] === 7) && (b[3] === 4 || b[3] === 6 || b[3] === 8); }

/** Returns { names, entries: Map(lowercase name -> entry), has(name), read(name) -> Buffer|null } */
function openZip(input, { maxEntryBytes = 512 * 1048576 } = {}) {
  const b = Buffer.isBuffer(input) ? input : Buffer.from(input.buffer, input.byteOffset, input.byteLength);
  if (b.length < 22) throw new ZipError('The file is too short to be a ZIP package.');
  let eocd = -1;
  for (let i = b.length - 22; i >= Math.max(0, b.length - 22 - 65535); i--) {
    if (b.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new ZipError('The ZIP package has no central directory: the file is cut off or damaged.');
  let count = b.readUInt16LE(eocd + 10);
  let cdSize = b.readUInt32LE(eocd + 12);
  let cdOff = b.readUInt32LE(eocd + 16);
  if (cdOff === 0xffffffff || count === 0xffff) {
    const loc = eocd - 20;
    if (loc >= 0 && b.readUInt32LE(loc) === 0x07064b50) {
      const z64 = Number(b.readBigUInt64LE(loc + 8));
      if (z64 + 56 <= b.length && b.readUInt32LE(z64) === 0x06064b50) {
        count = Number(b.readBigUInt64LE(z64 + 32));
        cdSize = Number(b.readBigUInt64LE(z64 + 40));
        cdOff = Number(b.readBigUInt64LE(z64 + 48));
      }
    }
  }
  if (cdOff + cdSize > b.length) throw new ZipError('The ZIP central directory points past the end of the file: the file is cut off.');
  const entries = new Map();
  const names = [];
  let p = cdOff;
  for (let i = 0; i < count && p + 46 <= b.length; i++) {
    if (b.readUInt32LE(p) !== 0x02014b50) throw new ZipError('The ZIP central directory is damaged.');
    const flags = b.readUInt16LE(p + 8);
    const method = b.readUInt16LE(p + 10);
    let csize = b.readUInt32LE(p + 20);
    let usize = b.readUInt32LE(p + 24);
    const nlen = b.readUInt16LE(p + 28);
    const xlen = b.readUInt16LE(p + 30);
    const clen = b.readUInt16LE(p + 32);
    let lho = b.readUInt32LE(p + 42);
    const name = b.toString(flags & 0x800 ? 'utf8' : 'latin1', p + 46, p + 46 + nlen);
    // zip64 extra field
    let x = p + 46 + nlen;
    const xEnd = x + xlen;
    while (x + 4 <= xEnd) {
      const id = b.readUInt16LE(x);
      const sz = b.readUInt16LE(x + 2);
      if (id === 0x0001) {
        let q = x + 4;
        if (usize === 0xffffffff && q + 8 <= x + 4 + sz) { usize = Number(b.readBigUInt64LE(q)); q += 8; }
        if (csize === 0xffffffff && q + 8 <= x + 4 + sz) { csize = Number(b.readBigUInt64LE(q)); q += 8; }
        if (lho === 0xffffffff && q + 8 <= x + 4 + sz) { lho = Number(b.readBigUInt64LE(q)); q += 8; }
      }
      x += 4 + sz;
    }
    const e = { name, method, csize, usize, lho, encrypted: !!(flags & 1) };
    entries.set(name.replace(/^\/+/, '').toLowerCase(), e);
    names.push(name);
    p += 46 + nlen + xlen + clen;
  }
  function get(name) { return entries.get(String(name).replace(/^\/+/, '').toLowerCase()) || null; }
  function read(name) {
    const e = typeof name === 'string' ? get(name) : name;
    if (!e) return null;
    if (e.encrypted) throw new ZipError(`The part ${e.name} is encrypted.`);
    if (e.lho + 30 > b.length || b.readUInt32LE(e.lho) !== 0x04034b50) throw new ZipError(`The part ${e.name} points to a damaged local header.`);
    const start = e.lho + 30 + b.readUInt16LE(e.lho + 26) + b.readUInt16LE(e.lho + 28);
    if (start + e.csize > b.length) throw new ZipError(`The part ${e.name} is cut off.`);
    const data = b.subarray(start, start + e.csize);
    if (e.usize > maxEntryBytes) throw new ZipError(`The part ${e.name} is ${Math.round(e.usize / 1048576)} MB unpacked, too large to read.`);
    if (e.method === 0) return Buffer.from(data);
    if (e.method === 8) {
      try { return zlib.inflateRawSync(data, { maxOutputLength: Math.max(1, Math.min(maxEntryBytes, e.usize || maxEntryBytes)) }); } catch (err) { throw new ZipError(`The part ${e.name} could not be unpacked: ${err.message}`); }
    }
    throw new ZipError(`The part ${e.name} uses ZIP method ${e.method}, which this viewer does not read.`);
  }
  return { names, entries, get, has: (n) => !!get(n), read };
}

module.exports = { openZip, isZip, ZipError };
