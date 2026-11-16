'use strict';
// Webview HTML for the macro viewer and the Pro tool panels. No vscode import, so tests can check the CSP: no remote
// origin, no connect-src, scripts only with the page nonce, styles from the extension's own folder.
const crypto = require('crypto');
const { footerHtml, FOOTER_CSS } = require('./support');

const PREFIX = 'vbaViewer';
const attr = (s) => String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

function csp(cspSource, n) {
  return ["default-src 'none'", `style-src ${cspSource} 'nonce-${n}'`, `script-src 'nonce-${n}'`, `img-src ${cspSource} data:`, `font-src ${cspSource}`].join('; ');
}

function head(p, n, footer) {
  return `<meta charset="UTF-8">
<meta http-equiv="Content-Security-Policy" content="${csp(p.cspSource, n)}">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="${attr(p.styleUri)}">
${footer ? `<style nonce="${n}">${FOOTER_CSS}</style>` : ''}
${typeof p.head === 'function' ? p.head(n) : ''}
<title>${attr(p.title)}</title>`;
}

/** p: { cspSource, scriptUri, styleUri, title, showSupportLinks, head?(nonce), nonce? } */
function editorHtml(p) {
  const n = p.nonce || crypto.randomBytes(18).toString('base64url');
  const footer = footerHtml(PREFIX, p.showSupportLinks !== false, '');
  return `<!DOCTYPE html>
<html lang="en">
<head>
${head(p, n, footer)}
</head>
<body>
<div id="toolbar" role="toolbar" aria-label="VBA project">
  <span id="chip" class="chip" title="Project name, modules and host"></span>
  <input id="search" type="search" placeholder="Search all modules" aria-label="Search all modules" spellcheck="false">
  <label class="tog" title="Show the hidden Attribute lines"><input type="checkbox" id="attrs"> Attributes</label>
  <button class="tb" id="copy" title="Copy the source of this module">Copy</button>
  <button class="tb" id="asText" title="Open this module as read-only text in an editor tab">Open as text</button>
</div>
<div id="safe" class="safe">Read-only: the macros are read as text and never run. Nothing is written to the file.</div>
<div id="err" role="alert" hidden></div>
<div id="warn" hidden><details><summary id="warnSum"></summary><ul id="warnList"></ul></details></div>
<div id="empty" hidden></div>
<div id="main" hidden>
  <nav id="list" aria-label="Modules"></nav>
  <section id="center">
    <div id="modhead"></div>
    <div id="code" tabindex="0" aria-label="Module source"></div>
  </section>
  <aside id="info" aria-label="Module details"></aside>
</div>
<div id="footer">${footer}</div>
<script nonce="${n}" src="${attr(p.scriptUri)}"></script>
</body>
</html>`;
}

/** Panel of a Pro tool. p: { cspSource, scriptUri, styleUri, title, head?(nonce), showSupportLinks, nonce? } */
function panelHtml(p) {
  const n = p.nonce || crypto.randomBytes(18).toString('base64url');
  const footer = footerHtml(PREFIX, p.showSupportLinks !== false, '');
  return `<!DOCTYPE html>
<html lang="en">
<head>
${head(p, n, footer)}
</head>
<body class="panel">
<div id="toolbar" role="toolbar"><strong id="title"></strong><span id="actions"></span></div>
<div id="app"><p class="muted">Working...</p></div>
<div id="footer">${footer}</div>
<script nonce="${n}" src="${attr(p.scriptUri)}"></script>
</body>
</html>`;
}

module.exports = { editorHtml, panelHtml, csp };
