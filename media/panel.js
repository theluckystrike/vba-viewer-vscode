// Panels of the Zovo Pro tools: Export all modules (what was written), Macro security scan (severity table, save as
// Markdown or JSON) and Compare macros (module table, a click opens the diff).
(function () {
  'use strict';
  const vscode = acquireVsCodeApi();
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s === undefined || s === null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const num = (n) => Number(n || 0).toLocaleString('en-US');
  const app = $('app');
  const TYPE = { document: 'document', standard: 'standard', class: 'class', form: 'form' };

  function exportView(m) {
    $('title').textContent = `Export all modules: ${m.file}`;
    $('actions').innerHTML = '<button class="tb" id="reveal">Show the folder</button>';
    let h = `<p class="msg" id="sum">${esc(m.summary)}</p>`;
    h += `<p class="muted small">Encoding: ${esc(m.encoding)}. ${esc(m.note || '')} A list of the files with their SHA-256 is in .vba-export.json in the folder; the next export compares against what is on disk and only rewrites files whose content changed.</p>`;
    h += '<table class="grid rep" id="rows"><thead><tr><th>File</th><th>Module</th><th>Type</th><th>Size</th><th>Status</th></tr></thead><tbody>'
      + m.rows.map((r) => `<tr data-file="${esc(r.file)}" class="st-${r.status}"><td><button class="lnk of" data-file="${esc(r.file)}">${esc(r.file)}</button></td><td>${esc(r.module)}</td><td>${esc(TYPE[r.type] || r.type)}</td><td class="r">${num(r.bytes)} B</td><td>${esc(r.status)}</td></tr>`).join('')
      + '</tbody></table>';
    if (m.orphans && m.orphans.length) h += `<p class="muted" id="orphans">No longer in the file, left on disk: ${m.orphans.map(esc).join(', ')}.</p>`;
    h += '<p class="muted small">Form modules are written as .frm with the form properties and code; the binary .frx layout file is not written.</p>';
    app.innerHTML = h;
  }

  function scanView(m) {
    const r = m.report;
    $('title').textContent = `Macro security scan: ${m.file}`;
    $('actions').innerHTML = '<button class="tb" id="saveMd" data-format="md">Save Markdown report</button><button class="tb" id="saveJson" data-format="json">Save JSON report</button>';
    const lvl = r.level === 'none' ? 'No findings' : `${r.level[0].toUpperCase()}${r.level.slice(1)} risk`;
    let h = `<p class="verdict lv-${r.level}" id="verdict"><strong>${esc(lvl)}.</strong> ${esc(r.summary)}</p>`;
    h += `<p id="counts" class="small">${r.counts.high} high · ${r.counts.medium} medium · ${r.counts.low} low · ${r.counts.info} info · ${num(r.modules)} module${r.modules === 1 ? '' : 's'}, ${num(r.lines)} code lines${r.xlmSheets ? `, ${r.xlmSheets} Excel 4.0 macro sheet${r.xlmSheets === 1 ? '' : 's'}` : ''}</p>`;
    if (r.findings.length) {
      h += '<table class="grid rep" id="rows"><thead><tr><th>Severity</th><th>Type</th><th>Keyword</th><th>Description</th><th>Count</th><th>Where</th></tr></thead><tbody>'
        + r.findings.map((f) => `<tr class="sev-${f.severity}"><td class="sv">${esc(f.severity)}</td><td>${esc(f.type)}</td><td class="mono">${esc(f.keyword)}</td><td>${esc(f.description)}${f.decoded ? `<div class="mono dec">decoded: ${esc(f.decoded)}</div>` : ''}</td><td class="r">${f.count}</td><td class="mono small">${esc(f.places.filter((p) => p.module).slice(0, 4).map((p) => (p.line ? `${p.module}:${p.line}` : p.module)).join(', '))}${f.places.length > 4 ? ', ...' : ''}</td></tr>`).join('')
        + '</tbody></table>';
    } else h += '<p class="ok" id="clean">Nothing to report.</p>';
    h += '<p class="muted small">The macros were read as text; nothing was run. Findings are signs to check, not proof that the file is harmful or safe.</p><div id="savedMsg"></div>';
    app.innerHTML = h;
  }

  function compareView(m) {
    $('title').textContent = `Compare macros: ${m.a} / ${m.b}`;
    $('actions').innerHTML = '';
    const s = m.sum;
    let h = `<p class="msg" id="sum">${s.changed} changed, ${s.onlyA} only in ${esc(m.a)}, ${s.onlyB} only in ${esc(m.b)}, ${s.same} the same${s.attributes ? `, ${s.attributes} with changed attributes only` : ''}. ${num(s.removed)} line${s.removed === 1 ? '' : 's'} removed and ${num(s.added)} added in all.</p>`;
    const label = { same: 'same', changed: 'changed', attributes: 'attributes only', onlyA: `only in ${m.a}`, onlyB: `only in ${m.b}` };
    h += '<table class="grid rep" id="rows"><thead><tr><th>Module</th><th>Status</th><th>Lines</th><th>Removed</th><th>Added</th><th></th></tr></thead><tbody>'
      + m.rows.map((r) => `<tr class="cs-${r.status}" data-name="${esc(r.name)}"><td>${esc(r.name)}</td><td>${esc(label[r.status])}</td><td class="r">${r.linesA} / ${r.linesB}</td><td class="r">${r.removed ? `-${num(r.removed)}` : ''}</td><td class="r">${r.added ? `+${num(r.added)}` : ''}</td><td><button class="lnk df" data-name="${esc(r.name)}">${r.status === 'onlyA' || r.status === 'onlyB' ? 'Open' : 'Diff'}</button></td></tr>`).join('')
      + '</tbody></table>';
    h += '<p class="muted small">Diff opens VS Code\'s side-by-side view of the two modules (read-only).</p>';
    app.innerHTML = h;
  }

  document.addEventListener('click', (e) => {
    const t = e.target;
    if (!t || !t.closest) return;
    const fmt = t.closest('[data-format]');
    if (fmt) { vscode.postMessage({ type: 'save', format: fmt.dataset.format }); return; }
    const of = t.closest('.of');
    if (of) { vscode.postMessage({ type: 'openFile', file: of.dataset.file }); return; }
    const df = t.closest('.df');
    if (df) { vscode.postMessage({ type: 'diff', name: df.dataset.name }); return; }
    if (t.id === 'reveal') vscode.postMessage({ type: 'reveal' });
  });

  window.addEventListener('message', (e) => {
    const m = e.data;
    if (!m || typeof m !== 'object') return;
    if (m.type === 'export') exportView(m);
    else if (m.type === 'scan') scanView(m);
    else if (m.type === 'compare') compareView(m);
    else if (m.type === 'saved') { const d = $('savedMsg'); if (d) d.innerHTML = `<p class="ok" id="saved">Saved ${esc(m.file)}.</p>`; }
    else if (m.type === 'error') { const d = document.createElement('p'); d.className = 'bad'; d.textContent = m.message; app.prepend(d); }
  });
  vscode.postMessage({ type: 'ready' });
}());
