// The macro viewer: module list, source with line numbers and VBA colors, search across all modules, module details,
// project overview with references, and Excel 4.0 macro sheets. Read-only; the data comes from the extension host.
(function () {
  'use strict';
  const vscode = acquireVsCodeApi();
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s === undefined || s === null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const num = (n) => Number(n || 0).toLocaleString('en-US');
  const bytes = (n) => (n >= 1048576 ? `${(n / 1048576).toFixed(1)} MB` : n >= 1024 ? `${(n / 1024).toFixed(1)} KB` : `${num(n)} bytes`);

  let data = null;
  let sel = { kind: 'module', idx: -1 };
  let curLine = 0;
  let showAttrs = false;
  let query = '';

  const ATTR_RE = /^\s*Attribute\s+[\w.]+\s*=/i;
  const KW = new Set(('addressof and as attribute base binary byref byval call case compare const database debug declare defbool defbyte defcur defdate '
    + 'defdbl defint deflng deflnglng deflngptr defobj defsng defstr defvar dim do each else elseif empty end enum eqv erase error event exit '
    + 'explicit false for friend function get global gosub goto if imp implements in is let lib like loop lset me mod module new next not nothing '
    + 'null on option optional or paramarray preserve print private property ptrsafe public raiseevent redim rem resume return rset select set '
    + 'static step stop sub text then to true type typeof until wend while with withevents xor open close output input append random access read '
    + 'write lock shared unload load line seek get put').split(' '));
  const TYPES = new Set('boolean byte currency date decimal double integer long longlong longptr object single string variant collection'.split(' '));
  const TYPE_ORDER = ['document', 'standard', 'class', 'form'];
  const GROUP = { document: 'Document modules', standard: 'Standard modules', class: 'Class modules', form: 'Forms' };

  /** VBA syntax colors for one line (HTML). */
  function hl(line) {
    if (ATTR_RE.test(line)) return `<span class="t-at">${esc(line)}</span>`;
    let out = '';
    let i = 0;
    const n = line.length;
    let stmtStart = true;
    while (i < n) {
      const c = line[i];
      if (c === '"') {
        let j = i + 1;
        while (j < n) { if (line[j] === '"') { if (line[j + 1] === '"') { j += 2; continue; } break; } j++; }
        out += `<span class="t-s">${esc(line.slice(i, j + 1))}</span>`;
        i = j + 1; stmtStart = false; continue;
      }
      if (c === "'") { out += `<span class="t-c">${esc(line.slice(i))}</span>`; break; }
      if (/\s/.test(c)) { out += c; i++; continue; }
      if (c === ':') { out += c; i++; stmtStart = true; continue; }
      if (c === '#' && stmtStart && /^#(If|Else|ElseIf|End If|Const)\b/i.test(line.slice(i))) {
        const m = /^#(?:End If|ElseIf|Else|If|Const)/i.exec(line.slice(i))[0];
        out += `<span class="t-p">${esc(m)}</span>`; i += m.length; stmtStart = false; continue;
      }
      if (c === '&' && /^&[HhOo][0-9A-Fa-f]+&?/.test(line.slice(i))) { const m = /^&[HhOo][0-9A-Fa-f]+&?/.exec(line.slice(i))[0]; out += `<span class="t-n">${esc(m)}</span>`; i += m.length; stmtStart = false; continue; }
      if (/\d/.test(c) && !/\w/.test(line[i - 1] || '')) { const m = /^\d+(?:\.\d+)?(?:[eE][-+]?\d+)?[%&!#@^]?/.exec(line.slice(i))[0]; out += `<span class="t-n">${esc(m)}</span>`; i += m.length; stmtStart = false; continue; }
      if (/[\p{L}_]/u.test(c)) {
        const m = /^[\p{L}_][\p{L}\p{N}_]*[$%&!#@]?/u.exec(line.slice(i))[0];
        const lw = m.toLowerCase().replace(/[$%&!#@]$/, '');
        if (lw === 'rem' && stmtStart) { out += `<span class="t-c">${esc(line.slice(i))}</span>`; break; }
        const prev = line[i - 1] === '.';
        if (!prev && KW.has(lw)) out += `<span class="t-k">${esc(m)}</span>`;
        else if (!prev && TYPES.has(lw)) out += `<span class="t-t">${esc(m)}</span>`;
        else out += esc(m);
        i += m.length; stmtStart = false; continue;
      }
      out += esc(c); i++; stmtStart = false;
    }
    return out;
  }

  /** Lines shown for a module: all of them, or without the Attribute lines (as the VBA editor shows it). */
  function visibleLines(m) {
    const all = String(m.source || '').split(/\r\n|\n|\r/);
    if (all.length && all[all.length - 1] === '') all.pop();
    return showAttrs ? all : all.filter((l) => !ATTR_RE.test(l));
  }

  /** Line number in the whole source (Attribute lines included) of a line number in the view. */
  function fullLine(m, vis) {
    if (showAttrs) return vis;
    const all = String(m.source || '').split(/\r\n|\n|\r/);
    let k = 0;
    for (let i = 0; i < all.length; i++) { if (ATTR_RE.test(all[i])) continue; k++; if (k === vis) return i + 1; }
    return 1;
  }

  function chipText() {
    const p = data.project || {};
    const n = data.modules.length;
    const parts = [p.name || data.name];
    parts.push(`${n} module${n === 1 ? '' : 's'}`);
    if (data.xlm && data.xlm.length) parts.push(`${data.xlm.length} Excel 4.0 macro sheet${data.xlm.length === 1 ? '' : 's'}`);
    if (data.host) parts.push(data.host);
    return parts.join(' · ');
  }

  // ---------- render ----------
  function render() {
    $('err').hidden = true; $('empty').hidden = true; $('main').hidden = true; $('warn').hidden = true;
    if (data.error) {
      $('err').hidden = false;
      $('err').innerHTML = `<p><strong>${esc(data.name)}</strong> could not be read.</p><p id="errText">${esc(data.error)}</p><p class="muted">Nothing was changed in the file.${/\.(xlsm|xlsb|xls|xltm|xlam|docm|dotm|doc|pptm|ppt)$/i.test(data.name) ? ' <button class="tb" id="errDefault">Reopen with the default editor</button>' : ''}</p>`;
      $('chip').textContent = data.name;
      for (const id of ['copy', 'asText', 'search', 'attrs']) $(id).disabled = true;
      vscode.postMessage({ type: 'failed' });
      return;
    }
    $('chip').textContent = chipText();
    if (data.warnings && data.warnings.length) {
      $('warn').hidden = false;
      $('warnSum').textContent = `${data.warnings.length} note${data.warnings.length === 1 ? '' : 's'} while reading the file`;
      $('warnList').innerHTML = data.warnings.map((w) => `<li>${esc(w)}</li>`).join('');
    }
    if (!data.hasVba && !(data.xlm && data.xlm.length)) {
      $('empty').hidden = false;
      for (const id of ['copy', 'asText', 'search', 'attrs']) $(id).disabled = true;
      $('empty').innerHTML = `<h2 id="emptyTitle">No VBA macros in this file</h2><p id="emptyNote">${esc(data.note || 'This file has no VBA project.')}</p>`
        + '<p class="muted">Office keeps VBA macros in a VBA project inside the file: vbaProject.bin in .xlsm, .xlsb, .xltm, .xlam, .docm, .dotm and .pptm packages, and a VBA storage in .xls, .doc and .ppt files. When a file is saved without macros, there is nothing to show.</p>'
        + `<p class="muted">File: ${esc(data.name)}, ${bytes(data.size)}${data.container ? `, ${data.container === 'ooxml' ? 'Office Open XML package' : 'OLE compound file'}` : ''}.</p>`;
      vscode.postMessage({ type: 'shown' });
      return;
    }
    for (const id of ['copy', 'asText', 'search', 'attrs']) $(id).disabled = !data.modules.length && id !== 'search';
    $('main').hidden = false;
    if (sel.idx < 0 || sel.kind === 'module' && !data.modules[sel.idx]) {
      const first = data.modules.findIndex((m) => m.lines > 0 && m.type !== 'document');
      const any = data.modules.findIndex((m) => m.lines > 0);
      sel = data.modules.length ? { kind: 'module', idx: first >= 0 ? first : any >= 0 ? any : 0 } : { kind: 'xlm', idx: 0 };
    }
    renderList();
    renderCenter();
    vscode.postMessage({ type: 'shown' });
  }

  function renderList() {
    let h = `<div class="grp"><button class="item proj${sel.kind === 'project' ? ' sel' : ''}" data-kind="project">Project overview</button></div>`;
    for (const t of TYPE_ORDER) {
      const ms = data.modules.map((m, i) => ({ m, i })).filter((x) => x.m.type === t);
      if (!ms.length) continue;
      h += `<div class="grp"><div class="gh">${GROUP[t]} <span class="muted">(${ms.length})</span></div>`
        + ms.map(({ m, i }) => `<button class="item mod${sel.kind === 'module' && sel.idx === i ? ' sel' : ''}" data-kind="module" data-idx="${i}" title="${esc(m.typeLabel)}: ${esc(m.name)}"><span class="nm">${esc(m.name)}</span>${m.error ? '<span class="bad">damaged</span>' : `<span class="ln">${num(m.lines)}</span>`}</button>`).join('')
        + '</div>';
    }
    if (data.xlm && data.xlm.length) {
      h += `<div class="grp"><div class="gh">Excel 4.0 macro sheets <span class="muted">(${data.xlm.length})</span></div>`
        + data.xlm.map((s, i) => `<button class="item xlm${sel.kind === 'xlm' && sel.idx === i ? ' sel' : ''}" data-kind="xlm" data-idx="${i}"><span class="nm">${esc(s.name)}</span><span class="ln">${esc(s.state)}</span></button>`).join('')
        + '</div>';
    }
    $('list').innerHTML = h;
  }

  function renderCenter() {
    if (query) { renderSearch(); return; }
    if (sel.kind === 'project') { renderProject(); return; }
    if (sel.kind === 'xlm') { renderXlm(data.xlm[sel.idx]); return; }
    const m = data.modules[sel.idx];
    if (!m) { $('modhead').innerHTML = ''; $('code').innerHTML = ''; $('info').innerHTML = ''; return; }
    const lines = visibleLines(m);
    $('modhead').innerHTML = `<strong id="modName">${esc(m.name)}</strong> <span class="muted">${esc(m.typeLabel)} · ${num(lines.length)} line${lines.length === 1 ? '' : 's'}</span>`;
    let body;
    if (m.error) body = `<p class="bad" id="modErr">${esc(m.error)}</p>`;
    else if (!lines.length || lines.every((l) => !l.trim())) body = `<p class="muted" id="modEmpty">This module has no code${showAttrs ? '' : ' (only hidden Attribute lines; tick Attributes to see them)'}.</p>`;
    else body = `<table class="src" id="srcTable"><tbody>${lines.map((l, i) => `<tr data-line="${i + 1}"${i + 1 === curLine ? ' class="cur"' : ''}><td class="no">${i + 1}</td><td class="tx">${hl(l) || ' '}</td></tr>`).join('')}</tbody></table>`;
    $('code').innerHTML = body;
    renderInfo(m);
    if (curLine) { const r = document.querySelector(`#srcTable tr[data-line="${curLine}"]`); if (r) r.scrollIntoView({ block: 'center' }); } else $('code').scrollTop = 0;
  }

  function renderInfo(m) {
    const row = (k, v, id) => `<tr${id ? ` id="${id}"` : ''}><th>${esc(k)}</th><td>${v}</td></tr>`;
    let h = '<h3>Module details</h3><table class="kv" id="kv"><tbody>';
    h += row('Name', esc(m.name), 'kvName');
    h += row('Type', esc(m.typeLabel), 'kvType');
    h += row('Stream', `VBA/${esc(m.streamName)}`);
    h += row('Source', `${bytes(m.size)}, ${num(m.lines)} code line${m.lines === 1 ? '' : 's'}`, 'kvSize');
    h += row('Stored', `${bytes(m.compressedSize)} compressed${m.offset ? ` (source at byte ${num(m.offset)}, after the compiled code)` : ''}`);
    if (m.docString) h += row('Description', esc(m.docString));
    const flags = [m.private ? 'private' : '', m.readOnly ? 'read-only' : '', m.hasFrame ? 'form layout stored' : ''].filter(Boolean);
    if (flags.length) h += row('Flags', esc(flags.join(', ')));
    h += '</tbody></table>';
    h += `<h3>Procedures <span class="muted">(${m.procedures.length})</span></h3>`;
    h += m.procedures.length ? `<ul class="procs" id="procs">${m.procedures.map((p) => `<li><button class="lnk proc" data-name="${esc(p.name)}">${esc(p.name)}</button> <span class="muted">${esc(p.kind)}</span></li>`).join('')}</ul>` : '<p class="muted">None.</p>';
    h += `<h3>Attributes <span class="muted">(${m.attributes.length})</span></h3>`;
    h += m.attributes.length ? `<ul class="attrs" id="attrList">${m.attributes.map((a) => `<li><code>${esc(a.replace(/^Attribute\s+/i, ''))}</code></li>`).join('')}</ul>` : '<p class="muted">None.</p>';
    $('info').innerHTML = h;
  }

  function renderProject() {
    const p = data.project || {};
    $('modhead').innerHTML = `<strong id="modName">Project overview</strong> <span class="muted">${esc(data.name)}</span>`;
    const row = (k, v) => `<tr><th>${esc(k)}</th><td>${v}</td></tr>`;
    let h = '<table class="kv" id="projKv"><tbody>';
    h += row('File', `${esc(data.name)} (${bytes(data.size)})`);
    h += row('Container', esc(data.container === 'ooxml' ? 'Office Open XML package (ZIP)' : data.container === 'vbaProject' ? 'VBA project file (vbaProject.bin)' : 'OLE compound file'));
    if (data.host) h += row('Office application', esc(data.host));
    if (data.vbaPath) h += row('VBA project location', esc(data.vbaPath));
    if (p.name) h += row('Project name', esc(p.name));
    if (p.docString) h += row('Description', esc(p.docString));
    if (p.codePage) h += row('Code page', `${p.codePage} (${esc(p.codePageLabel)})`);
    if (p.sysKindLabel) h += row('Platform', esc(p.sysKindLabel));
    if (p.versionMajor || p.versionMinor) h += row('Project version', `${p.versionMajor}.${p.versionMinor}`);
    if (p.constants) h += row('Conditional constants', esc(p.constants));
    const counts = TYPE_ORDER.map((t) => [t, data.modules.filter((m) => m.type === t).length]).filter((x) => x[1]).map(([t, n]) => `${n} ${GROUP[t].toLowerCase()}`);
    h += row('Modules', esc(counts.join(', ') || 'none'));
    h += row('VBA password', data.protected ? 'Yes. It only hides the code in Office; the source is still stored in the file and shows here.' : 'No');
    h += '</tbody></table>';
    if (data.references && data.references.length) {
      h += `<h3>References (${data.references.length})</h3><table class="grid" id="refs"><thead><tr><th>Name</th><th>Kind</th><th>Library</th></tr></thead><tbody>`
        + data.references.map((r) => `<tr><td>${esc(r.name)}</td><td>${esc(r.kind)}</td><td class="mono">${esc(r.libid)}</td></tr>`).join('') + '</tbody></table>';
    }
    if (data.xlmNames && data.xlmNames.length) h += `<h3>Excel 4.0 auto-run names</h3><ul>${data.xlmNames.map((x) => `<li>${esc(x.name)}${x.ref ? ` = ${esc(x.ref)}` : ''}</li>`).join('')}</ul>`;
    $('code').innerHTML = `<div class="pad">${h}</div>`;
    $('info').innerHTML = '<h3>About this view</h3><p class="muted">The macros were read from the file as text. None of them ran, and the file was not changed.</p>';
  }

  function renderXlm(s) {
    if (!s) return;
    $('modhead').innerHTML = `<strong id="modName">${esc(s.name)}</strong> <span class="muted">${esc(s.kind)} · ${esc(s.state)}</span>`;
    let h = '';
    if (s.formulas && s.formulas.length) {
      h = `<table class="grid" id="xlmTable"><thead><tr><th>Cell</th><th>Formula</th></tr></thead><tbody>${s.formulas.map((f) => `<tr><td class="mono">${esc(f.cell)}</td><td class="mono">${esc(f.formula)}</td></tr>`).join('')}</tbody></table>`;
    } else h = `<p class="muted" id="xlmNone">${s.part && /\.xml$/i.test(s.part) ? 'This macro sheet has no formulas.' : 'The formulas of this macro sheet are stored in binary form and are not listed; the sheet name and visibility are.'}</p>`;
    $('code').innerHTML = `<div class="pad">${h}</div>`;
    $('info').innerHTML = `<h3>Macro sheet</h3><table class="kv"><tbody><tr><th>Name</th><td>${esc(s.name)}</td></tr><tr><th>Visibility</th><td>${esc(s.state)}</td></tr>${s.part ? `<tr><th>Part</th><td>${esc(s.part)}</td></tr>` : ''}<tr><th>Formulas</th><td>${num((s.formulas || []).length)}</td></tr></tbody></table><p class="muted">Excel 4.0 macros are formulas such as EXEC or CALL on a macro sheet. They are shown as text; nothing is calculated.</p>`;
  }

  function renderSearch() {
    const q = query.toLowerCase();
    const hits = [];
    let total = 0;
    data.modules.forEach((m, mi) => {
      visibleLines(m).forEach((l, li) => {
        const at = l.toLowerCase().indexOf(q);
        if (at < 0) return;
        total++;
        if (hits.length < 1000) hits.push({ mi, line: li + 1, text: l, at });
      });
    });
    const mods = new Set(hits.map((h) => h.mi)).size;
    $('modhead').innerHTML = `<strong id="modName">Search</strong> <span class="muted" id="searchSum">${total ? `${num(total)} line${total === 1 ? '' : 's'} with "${esc(query)}" in ${mods} module${mods === 1 ? '' : 's'}${total > hits.length ? `, first ${hits.length} shown` : ''}` : `No line has "${esc(query)}"`}</span>`;
    let h = '<div id="results">';
    let lastM = -1;
    for (const x of hits) {
      if (x.mi !== lastM) { if (lastM >= 0) h += '</div>'; h += `<div class="rgrp"><div class="rh">${esc(data.modules[x.mi].name)}</div>`; lastM = x.mi; }
      const t = x.text;
      const s = Math.max(0, x.at - 60);
      const snip = `${s ? '...' : ''}${esc(t.slice(s, x.at))}<mark>${esc(t.slice(x.at, x.at + query.length))}</mark>${esc(t.slice(x.at + query.length, x.at + query.length + 120))}`;
      h += `<button class="hit" data-idx="${x.mi}" data-line="${x.line}"><span class="no">${x.line}</span><span class="tx">${snip}</span></button>`;
    }
    if (lastM >= 0) h += '</div>';
    h += '</div>';
    $('code').innerHTML = h;
  }

  // ---------- events ----------
  document.addEventListener('click', (e) => {
    const t = e.target.closest ? e.target : null;
    if (!t) return;
    const item = t.closest('.item');
    if (item) {
      query = ''; $('search').value = '';
      sel = { kind: item.dataset.kind, idx: Number(item.dataset.idx || 0) };
      curLine = 0;
      if (sel.kind === 'module' && data.modules[sel.idx]) vscode.postMessage({ type: 'select', module: data.modules[sel.idx].name });
      renderList(); renderCenter();
      return;
    }
    const hit = t.closest('.hit');
    if (hit) {
      sel = { kind: 'module', idx: Number(hit.dataset.idx) };
      curLine = Number(hit.dataset.line);
      query = ''; $('search').value = '';
      renderList(); renderCenter();
      return;
    }
    const proc = t.closest('.proc');
    if (proc) {
      const m = data.modules[sel.idx];
      const re = new RegExp(`^\\s*(?:(?:Public|Private|Friend)\\s+)?(?:Static\\s+)?(?:Sub|Function|Property\\s+(?:Get|Let|Set))\\s+${proc.dataset.name}(?![\\p{L}\\p{N}_])`, 'iu');
      const i = visibleLines(m).findIndex((l) => re.test(l));
      if (i >= 0) { curLine = i + 1; renderCenter(); }
      return;
    }
    const tr = t.closest('#srcTable tr');
    if (tr) {
      curLine = Number(tr.dataset.line);
      document.querySelectorAll('#srcTable tr.cur').forEach((r) => r.classList.remove('cur'));
      tr.classList.add('cur');
      return;
    }
    if (t.id === 'copy') {
      const m = sel.kind === 'module' && data.modules[sel.idx];
      if (m) { const ls = visibleLines(m); vscode.postMessage({ type: 'copy', text: ls.join('\r\n') + '\r\n', what: `${m.name} (${ls.length} lines)` }); }
      return;
    }
    if (t.id === 'asText') {
      const m = sel.kind === 'module' && data.modules[sel.idx];
      if (m) vscode.postMessage({ type: 'openText', module: m.name, line: fullLine(m, curLine || 1) });
      return;
    }
    if (t.id === 'errDefault') vscode.postMessage({ type: 'openDefault' });
  });
  let timer = null;
  $('search').addEventListener('input', () => {
    clearTimeout(timer);
    timer = setTimeout(() => { query = $('search').value.trim(); renderCenter(); }, 150);
  });
  $('search').addEventListener('keydown', (e) => { if (e.key === 'Escape') { $('search').value = ''; query = ''; renderCenter(); } });
  $('attrs').addEventListener('change', () => { showAttrs = $('attrs').checked; curLine = 0; if (data && !data.error) renderCenter(); });
  document.addEventListener('keydown', (e) => {
    if (e.key === '/' && document.activeElement !== $('search')) { e.preventDefault(); $('search').focus(); }
  });

  window.addEventListener('message', (e) => {
    const m = e.data;
    if (!m || m.type !== 'load') return;
    const prevName = data && data.modules && data.modules[sel.idx] ? data.modules[sel.idx].name : null;
    data = m;
    if (!data.modules) data.modules = [];
    showAttrs = !!m.showAttributes;
    $('attrs').checked = showAttrs;
    if (prevName) { const i = data.modules.findIndex((x) => x.name === prevName); sel = i >= 0 ? { kind: 'module', idx: i } : { kind: 'module', idx: -1 }; }
    render();
  });
  vscode.postMessage({ type: 'ready' });
}());
