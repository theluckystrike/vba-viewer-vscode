'use strict';
// Compare the macros of two files module by module: same, changed (lines added and removed, from the shortest edit
// script of Myers' O(ND) algorithm), only in the first file, only in the second. The diff view itself is VS Code's.
const { analyzeSource } = require('./vba');

/** Number of lines removed and added between two line arrays. */
function lineDelta(a, b, maxD = 6000) {
  const n = a.length;
  const m = b.length;
  const max = Math.min(n + m, maxD);
  const off = max + 1;
  const v = new Int32Array(2 * max + 3);
  for (let d = 0; d <= max; d++) {
    for (let k = -d; k <= d; k += 2) {
      let x = k === -d || (k !== d && v[off + k - 1] < v[off + k + 1]) ? v[off + k + 1] : v[off + k - 1] + 1;
      let y = x - k;
      while (x < n && y < m && a[x] === b[y]) { x++; y++; }
      v[off + k] = x;
      if (x >= n && y >= m) { const lcs = (n + m - d) / 2; return { removed: n - lcs, added: m - lcs, exact: true }; }
    }
  }
  // Very different modules: count lines that are not in the other file (multiset), an upper bound of the script.
  const count = new Map();
  for (const l of a) count.set(l, (count.get(l) || 0) + 1);
  let common = 0;
  for (const l of b) { const c = count.get(l); if (c) { common++; count.set(l, c - 1); } }
  return { removed: n - common, added: m - common, exact: false };
}

/** rows: [{ name, typeA, typeB, status: 'same'|'changed'|'attributes'|'onlyA'|'onlyB', added, removed, linesA, linesB }] */
function compareResults(a, b) {
  const byName = (r) => new Map((r.modules || []).map((m) => [m.name.toLowerCase(), m]));
  const A = byName(a);
  const B = byName(b);
  const names = [...new Set([...A.keys(), ...B.keys()])];
  const rows = [];
  for (const k of names) {
    const ma = A.get(k);
    const mb = B.get(k);
    if (ma && !mb) { rows.push({ name: ma.name, typeA: ma.type, typeB: '', status: 'onlyA', added: 0, removed: ma.lines, linesA: ma.lines, linesB: 0 }); continue; }
    if (!ma && mb) { rows.push({ name: mb.name, typeA: '', typeB: mb.type, status: 'onlyB', added: mb.lines, removed: 0, linesA: 0, linesB: mb.lines }); continue; }
    const ca = analyzeSource(ma.source).code;
    const cb = analyzeSource(mb.source).code;
    const norm = (s) => String(s).replace(/\r\n?/g, '\n');
    if (norm(ma.source) === norm(mb.source)) { rows.push({ name: ma.name, typeA: ma.type, typeB: mb.type, status: 'same', added: 0, removed: 0, linesA: ca.length, linesB: cb.length }); continue; }
    const d = lineDelta(ca, cb);
    rows.push({ name: ma.name, typeA: ma.type, typeB: mb.type, status: d.added || d.removed ? 'changed' : 'attributes', added: d.added, removed: d.removed, exact: d.exact, linesA: ca.length, linesB: cb.length });
  }
  const order = { changed: 0, attributes: 1, onlyA: 2, onlyB: 3, same: 4 };
  rows.sort((x, y) => (order[x.status] - order[y.status]) || x.name.localeCompare(y.name, 'en', { sensitivity: 'base' }));
  const sum = { same: 0, changed: 0, attributes: 0, onlyA: 0, onlyB: 0, added: 0, removed: 0 };
  for (const r of rows) { sum[r.status]++; sum.added += r.added; sum.removed += r.removed; }
  return { rows, sum };
}

module.exports = { compareResults, lineDelta };
