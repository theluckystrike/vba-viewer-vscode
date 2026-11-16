'use strict';
// PowerPoint 97-2003 (.ppt) keeps its VBA project as an embedded OLE storage inside the "PowerPoint Document"
// stream: an ExOleObjStg record (type 0x1011), zlib-compressed when its instance is 1 ([MS-PPT] 2.10.34). This walks
// the record tree, unpacks each such record and keeps the first one that is a compound file with a VBA storage.
const zlib = require('zlib');
const { isCfb, readCfb } = require('./cfb');

const EX_OLE_OBJ_STG = 0x1011;

function findPptVba(cfb) {
  const e = cfb.child(cfb.root, 'PowerPoint Document');
  if (!e || e.type !== 2) return null;
  const s = cfb.read(e);
  const stack = [[0, s.length, 0]];
  let steps = 0;
  while (stack.length) {
    const [start, end, depth] = stack.pop();
    let p = start;
    while (p + 8 <= end && steps++ < 2000000) {
      const verInst = s.readUInt16LE(p);
      const type = s.readUInt16LE(p + 2);
      const len = s.readUInt32LE(p + 4);
      const body = p + 8;
      if (len > end - body) break;
      const ver = verInst & 0x0f;
      const inst = verInst >> 4;
      if (ver === 0x0f && depth < 16) stack.push([body, body + len, depth + 1]);
      else if (type === EX_OLE_OBJ_STG && len > 4) {
        let bytes = null;
        try {
          bytes = inst === 1 ? zlib.inflateSync(s.subarray(body + 4, body + len), { maxOutputLength: 256 * 1048576 }) : s.subarray(body, body + len);
        } catch (err) { bytes = null; }
        if (bytes && isCfb(bytes)) {
          try {
            const inner = readCfb(bytes);
            const vba = inner.child(inner.root, 'VBA');
            if (vba && vba.type === 1 && inner.child(vba, 'dir')) return { bytes, offset: p };
          } catch (err) { /* not a usable storage, keep looking */ }
        }
      }
      p = body + len;
    }
  }
  return null;
}

module.exports = { findPptVba };
