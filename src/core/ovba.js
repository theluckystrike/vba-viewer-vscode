'use strict';
// MS-OVBA compression (section 2.4.1 of the public [MS-OVBA] specification), written for this extension: the
// decompression that every VBA module, dir stream and form uses, and the matching compression (used by the tests and
// by the sample builder). Unit tests check both against the examples in section 3.2 of the specification.

class OvbaError extends Error {}

class Grow {
  constructor(n) { this.b = new Uint8Array(Math.max(256, n | 0)); this.n = 0; }
  need(k) { if (this.n + k <= this.b.length) return; let m = this.b.length * 2; while (m < this.n + k) m *= 2; const nb = new Uint8Array(m); nb.set(this.b.subarray(0, this.n)); this.b = nb; }
  push(v) { this.need(1); this.b[this.n++] = v; }
  bytes() { return Buffer.from(this.b.buffer, 0, this.n); }
}

/** bit count of a CopyToken for a position `diff` bytes into the decompressed chunk (CopyToken Help, 2.4.1.3.19.1) */
function bitCount(diff) {
  let bc = 4;
  while ((1 << bc) < diff) bc++;
  return bc;
}

/**
 * Decompresses a CompressedContainer that starts at `start` in `buf`. Returns a Buffer.
 * Throws OvbaError on a bad signature byte, a bad chunk signature or a copy token that points before its chunk.
 */
function decompress(input, start = 0, end) {
  const buf = Buffer.isBuffer(input) ? input : Buffer.from(input.buffer, input.byteOffset, input.byteLength);
  const stop = end == null ? buf.length : Math.min(end, buf.length);
  if (start >= stop) throw new OvbaError('The compressed data is empty.');
  if (buf[start] !== 0x01) throw new OvbaError(`The compressed data does not start with the signature byte 0x01 (found 0x${buf[start].toString(16).padStart(2, '0')}).`);
  const out = new Grow((stop - start) * 3);
  let pos = start + 1;
  while (pos + 2 <= stop) {
    const hdr = buf.readUInt16LE(pos);
    const size = (hdr & 0x0fff) + 3;
    if (((hdr >> 12) & 0x07) !== 0x03) throw new OvbaError(`A compressed chunk at byte ${pos - start} has a bad signature (header 0x${hdr.toString(16)}).`);
    const compressed = (hdr & 0x8000) !== 0;
    const chunkEnd = Math.min(pos + size, stop);
    let p = pos + 2;
    const chunkStart = out.n;
    if (!compressed) {
      const n = Math.min(4096, stop - p);
      out.need(n);
      out.b.set(buf.subarray(p, p + n), out.n);
      out.n += n;
      pos = p + 4096;
      continue;
    }
    while (p < chunkEnd) {
      const flags = buf[p++];
      for (let bit = 0; bit < 8 && p < chunkEnd; bit++) {
        if (((flags >> bit) & 1) === 0) { out.push(buf[p++]); continue; }
        if (p + 2 > chunkEnd) throw new OvbaError('A copy token is cut off at the end of a compressed chunk.');
        const token = buf.readUInt16LE(p);
        p += 2;
        const diff = out.n - chunkStart;
        const bc = bitCount(diff);
        const lengthMask = 0xffff >> bc;
        const length = (token & lengthMask) + 3;
        const offset = (token >> (16 - bc)) + 1;
        if (offset > diff) throw new OvbaError('A copy token points before the start of its chunk; the data is damaged.');
        out.need(length);
        for (let k = 0; k < length; k++) { out.b[out.n] = out.b[out.n - offset]; out.n++; }
      }
    }
    pos = chunkEnd;
  }
  return out.bytes();
}

/** Compresses bytes into a CompressedContainer exactly as section 2.4.1.3.6 describes. */
function compress(input) {
  const src = Buffer.isBuffer(input) ? input : Buffer.from(typeof input === 'string' ? Buffer.from(input, 'latin1') : input);
  const out = new Grow(src.length + 16);
  out.push(0x01);
  let dc = 0;
  const dEndAll = src.length;
  while (dc < dEndAll) {
    const cStart = out.n;
    const dStart = dc;
    const cEnd = cStart + 4098;
    const dEnd = Math.min(dStart + 4096, dEndAll);
    out.need(4098);
    out.n = cStart + 2;
    while (dc < dEnd && out.n < cEnd) {
      // token sequence
      const flagAt = out.n;
      let flags = 0;
      out.n++;
      for (let i = 0; i < 8; i++) {
        if (!(dc < dEnd && out.n < cEnd)) continue;
        // matching (2.4.1.3.19.4)
        let best = 0;
        let bestCand = 0;
        for (let cand = dc - 1; cand >= dStart; cand--) {
          let c = cand;
          let dd = dc;
          let len = 0;
          while (dd < dEnd && src[dd] === src[c]) { len++; c++; dd++; }
          if (len > best) { best = len; bestCand = cand; }
        }
        let offset = 0;
        let length = 0;
        if (best >= 3) {
          const bc = bitCount(dc - dStart);
          const maxLen = (0xffff >> bc) + 3;
          length = Math.min(best, maxLen);
          offset = dc - bestCand;
        }
        if (offset) {
          if (out.n + 1 < cEnd) {
            const bc = bitCount(dc - dStart);
            const token = ((offset - 1) << (16 - bc)) | (length - 3);
            out.b[out.n] = token & 0xff;
            out.b[out.n + 1] = token >> 8;
            flags |= 1 << i;
            out.n += 2;
            dc += length;
          } else out.n = cEnd;
        } else if (out.n < cEnd) {
          out.b[out.n++] = src[dc++];
        } else out.n = cEnd;
      }
      out.b[flagAt] = flags;
    }
    let flag = 1;
    if (dc < dEnd) {
      // raw chunk
      out.n = cStart + 2;
      dc = dStart;
      let pad = 4096;
      for (let k = dStart; k <= dEnd - 1; k++) { out.b[out.n++] = src[k]; dc++; pad--; }
      for (let k = 0; k < pad; k++) out.b[out.n++] = 0;
      flag = 0;
    }
    const size = out.n - cStart;
    const hdr = ((size - 3) & 0x0fff) | 0x3000 | (flag << 15);
    out.b[cStart] = hdr & 0xff;
    out.b[cStart + 1] = hdr >> 8;
  }
  return out.bytes();
}

module.exports = { decompress, compress, OvbaError, bitCount };
