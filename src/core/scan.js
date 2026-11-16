'use strict';
// Macro security scan: reads the macro source as text and lists what deserves a look before anyone enables the
// macros: auto-run entry points, calls that start programs, download files or touch memory, hidden strings (hex,
// Base64, character codes), URLs, IP addresses and file names, and Excel 4.0 macro sheets. Nothing is run or
// evaluated. Findings are signs to check, not proof that a file is harmful.
const { analyzeSource } = require('./vba');

const SEV_RANK = { high: 3, medium: 2, low: 1, info: 0 };

const AUTORUN = [
  ['Workbook_Open', 'Runs when the workbook opens'], ['Auto_Open', 'Runs when the workbook or add-in opens'], ['AutoOpen', 'Runs when the document opens'],
  ['Document_Open', 'Runs when the document opens'], ['AutoExec', 'Runs when Word starts or the template loads'], ['AutoNew', 'Runs when a document is made from the template'],
  ['Document_New', 'Runs when a document is made from the template'], ['AutoClose', 'Runs when the document closes'], ['Document_Close', 'Runs when the document closes'],
  ['AutoExit', 'Runs when Word quits'], ['Auto_Close', 'Runs when the workbook closes'], ['Workbook_BeforeClose', 'Runs when the workbook closes'],
  ['Workbook_Activate', 'Runs when the workbook window gets focus'], ['Workbook_WindowActivate', 'Runs when a workbook window gets focus'],
  ['Document_ContentControlOnEnter', 'Runs when a content control gets focus'], ['App_WorkbookOpen', 'Runs when any workbook opens (application event)'],
];
// ActiveX control events that fire without a click when the document opens
const AUTORUN_SUFFIX = [['_Painted', 'An ActiveX control event that fires when the control is drawn'], ['_Layout', 'An ActiveX control event that fires when the control is laid out'],
  ['_GotFocus', 'An ActiveX control event that can fire when the document opens'], ['_Resize', 'An ActiveX control event that fires when the control is sized']];

// [regex on code without comments, keyword shown, description, severity]
const SUSPICIOUS = [
  [/\bShell\s*\(|^\s*Shell\s+["\w]|[=:]\s*Shell\s+["\w]/i, 'Shell', 'Starts another program', 'high'],
  [/WScript\.Shell/i, 'WScript.Shell', 'Windows Script Host object that runs commands', 'high'],
  [/Shell\.Application/i, 'Shell.Application', 'Windows shell object that can start programs', 'high'],
  [/\bShellExecute(?:A|W|Ex)?\b/i, 'ShellExecute', 'Windows API call that starts a program', 'high'],
  [/powershell/i, 'PowerShell', 'Mentions PowerShell, often used to run downloaded code', 'high'],
  [/\bcmd(?:\.exe)?\s*\/[ck]\b|\bcmd\.exe\b/i, 'cmd.exe', 'Runs a command prompt command', 'high'],
  [/\bURLDownloadToFile[AW]?\b/i, 'URLDownloadToFile', 'Downloads a file from the internet', 'high'],
  [/\b(?:mshta|regsvr32|rundll32|certutil|bitsadmin|schtasks|wmic)(?:\.exe)?\b/i, 'Windows tool', 'Calls a Windows tool that attackers often use to run or fetch code', 'high'],
  [/Win32_Process\b/i, 'Win32_Process', 'Starts a process through WMI', 'high'],
  [/\b(?:VirtualAlloc(?:Ex)?|RtlMoveMemory|CreateThread|WriteProcessMemory|VirtualProtect|EnumSystemLanguageGroupsW?|CallWindowProc[AW]?)\b/i, 'Memory API', 'Windows memory functions used to run machine code from a macro', 'high'],
  [/\bExecuteExcel4Macro\b/i, 'ExecuteExcel4Macro', 'Runs an Excel 4.0 macro formula from VBA', 'high'],
  [/\bVBProject\b|\bVBComponents\b|\bCodeModule\b/i, 'VBProject', 'Reads or changes VBA code (a macro that edits macros)', 'high'],
  [/\bDDE(?:Initiate|Execute)\b/i, 'DDE', 'Dynamic Data Exchange, can start programs', 'high'],
  [/\bCreateObject\s*\(/i, 'CreateObject', 'Creates a COM object (look at which one)', 'medium'],
  [/\bGetObject\s*\(/i, 'GetObject', 'Gets a COM object, for example WMI', 'medium'],
  [/\b(?:MSXML2\.(?:Server)?XMLHTTP|Microsoft\.XMLHTTP|WinHttp\.WinHttpRequest|MSXML2\.ServerXMLHTTP)/i, 'XMLHTTP', 'Sends or downloads data over the network', 'medium'],
  [/\bInternetOpen(?:Url)?[AW]?\b|\bInternetReadFile\b/i, 'WinINet', 'Windows API calls that download data', 'medium'],
  [/ADODB\.Stream/i, 'ADODB.Stream', 'Writes binary data to files', 'medium'],
  [/\.SaveToFile\b/i, 'SaveToFile', 'Saves data to a file', 'medium'],
  [/\bDeclare\s+(?:PtrSafe\s+)?(?:Function|Sub)\b[^\n]*\bLib\b/i, 'Declare Lib', 'Calls a Windows DLL function directly', 'medium'],
  [/\bCallByName\b/i, 'CallByName', 'Calls a method named in a string, which hides what runs', 'medium'],
  [/\bStrReverse\b/i, 'StrReverse', 'Reverses text, often used to hide strings', 'medium'],
  [/\bSendKeys\b/i, 'SendKeys', 'Types keys into the active window', 'medium'],
  [/\bReg(?:Write|Read|Delete)\b|\bGetSetting\b|\bSaveSetting\b/i, 'Registry', 'Reads or writes the Windows registry', 'medium'],
  [/^\s*Kill\s|\bKill\s+["\w]/i, 'Kill', 'Deletes files', 'medium'],
  [/(?<!Application)\.Run\b/i, 'Run', 'Runs a command or macro by name', 'low'],
  [/Scripting\.FileSystemObject/i, 'FileSystemObject', 'Reads and writes files', 'low'],
  [/\bCreateTextFile\b|\bOpen\b[^\n]*\bFor\s+(?:Output|Binary|Append)\b/i, 'Write file', 'Writes a file', 'low'],
  [/\bEnviron\$?\s*\(/i, 'Environ', 'Reads environment variables such as the user name or temp folder', 'low'],
  [/\bApplication\.Run\b/i, 'Application.Run', 'Runs a macro by name', 'low'],
];

const URL_RE = /\b(?:https?|ftp):\/\/[^\s"'<>)]+/gi;
const IP_RE = /\b(?:(?:25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)\.){3}(?:25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)\b/g;
const EXE_RE = /\b[\w.-]+\.(?:exe|dll|scr|pif|bat|cmd|ps1|vbs|vbe|jse|hta|wsf|lnk|jar)\b/gi;
const XLM_RISK = /\b(?:EXEC|CALL|REGISTER|RUN|FOPEN|FWRITE|FWRITELN|URLDownloadToFileA?|ALERT|HALT|FORMULA|GOTO|CHAR)\s*\(/i;

/** Splits one physical line into code (string contents kept), its string literals, and whether it ended in a comment. */
function splitLine(line) {
  const strings = [];
  let code = '';
  let i = 0;
  const n = line.length;
  // Rem comment at the start of a statement
  if (/^\s*Rem(\s|$)/i.test(line)) return { code: '', strings, comment: line };
  while (i < n) {
    const c = line[i];
    if (c === '"') {
      let j = i + 1;
      let s = '';
      while (j < n) {
        if (line[j] === '"') { if (line[j + 1] === '"') { s += '"'; j += 2; continue; } break; }
        s += line[j++];
      }
      strings.push(s);
      code += line.slice(i, Math.min(j + 1, n));
      i = j + 1;
      continue;
    }
    if (c === "'") return { code, strings, comment: line.slice(i) };
    if (/[:\s]/.test(c) && /^[:\s]\s*Rem(\s|$)/i.test(line.slice(i)) && /:\s*$/.test(code + c)) return { code, strings, comment: line.slice(i) };
    code += c;
    i++;
  }
  return { code, strings, comment: '' };
}

function printable(buf) {
  if (!buf.length) return 0;
  let k = 0;
  for (const c of buf) if ((c >= 0x20 && c < 0x7f) || c === 9 || c === 10 || c === 13) k++;
  return k / buf.length;
}
const preview = (s, n = 80) => (s.length > n ? `${s.slice(0, n)}...` : s);
const DECODED_RISK = /(?:powershell|cmd\.exe|https?:\/\/|\.exe\b|WScript|Shell|URLDownload|CreateObject|XMLHTTP|ADODB)/i;

/**
 * Scans the result of vba.extract(). Returns { findings: [{ type, keyword, description, severity, count, places: [{ module, line }] , decoded? }],
 * counts: { high, medium, low, info }, level, summary, modules, lines }.
 */
function scan(result, opts = {}) {
  const fileName = opts.fileName || '';
  const map = new Map();
  const add = (type, keyword, description, severity, module, line, extra) => {
    const k = `${type}\u0000${keyword}\u0000${extra && extra.decoded ? extra.decoded : ''}`;
    let f = map.get(k);
    if (!f) { f = { type, keyword, description, severity, count: 0, places: [], ...(extra || {}) }; map.set(k, f); }
    f.count++;
    if (f.places.length < 200 && module != null) f.places.push({ module, line });
  };
  let totalLines = 0;
  const autoNames = new Set(AUTORUN.map(([n]) => n.toLowerCase()));
  for (const m of result.modules || []) {
    const { code } = analyzeSource(m.source || '');
    totalLines += code.length;
    let chrCount = 0;
    let chrFirst = 0;
    code.forEach((raw, idx) => {
      const ln = idx + 1;
      const { code: c, strings } = splitLine(raw);
      if (!c.trim()) return;
      const sub = /^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function)\s+([\p{L}_][\p{L}\p{N}_]*)/iu.exec(c);
      if (sub) {
        const name = sub[1];
        const a = AUTORUN.find(([x]) => x.toLowerCase() === name.toLowerCase());
        if (a) add('Auto-run', a[0], a[1], 'high', m.name, ln);
        else if (/_/.test(name) && !autoNames.has(name.toLowerCase())) {
          const s = AUTORUN_SUFFIX.find(([x]) => name.toLowerCase().endsWith(x.toLowerCase()));
          if (s) add('Auto-run', name, s[1], 'medium', m.name, ln);
        }
      }
      for (const [re, kw, desc, sev] of SUSPICIOUS) if (re.test(c)) add('Suspicious', kw, desc, sev, m.name, ln);
      const chr = c.match(/\bChr[WB]?\$?\s*\(/gi);
      if (chr) { if (!chrCount) chrFirst = ln; chrCount += chr.length; }
      for (const s of strings) {
        for (const u of s.match(URL_RE) || []) add('IOC', u, 'URL', 'medium', m.name, ln);
        for (const ip of s.match(IP_RE) || []) if (!/^(?:0\.0\.0\.0|127\.0\.0\.1|255\.255\.255\.255)$/.test(ip) && !/\d\.\d+\.\d+\.\d+\.\d/.test(s)) add('IOC', ip, 'IP address', 'medium', m.name, ln);
        for (const x of s.match(EXE_RE) || []) add('IOC', x, 'Executable file name', 'medium', m.name, ln);
        const t = s.trim();
        if (t.length >= 16 && t.length % 2 === 0 && /^[0-9A-Fa-f]+$/.test(t) && !/^\d+$/.test(t)) {
          const dec = Buffer.from(t, 'hex');
          if (printable(dec) >= 0.85) { const d = dec.toString('latin1'); add('Hex string', preview(t, 40), 'Hex-encoded text', DECODED_RISK.test(d) ? 'high' : 'low', m.name, ln, { decoded: preview(d) }); }
        } else if (t.length >= 24 && t.length % 4 === 0 && /^[A-Za-z0-9+/]+={0,2}$/.test(t) && /[A-Z]/.test(t) && /[a-z]/.test(t) && /\d|\+|\//.test(t)) {
          const dec = Buffer.from(t, 'base64');
          if (dec.length >= 12 && printable(dec) >= 0.85) { const d = dec.toString('latin1'); add('Base64 string', preview(t, 40), 'Base64-encoded text', DECODED_RISK.test(d) ? 'high' : 'low', m.name, ln, { decoded: preview(d) }); }
        }
      }
    });
    if (chrCount >= 20) add('Obfuscation', 'Chr()', `Builds text from ${chrCount} character codes, a common way to hide strings`, chrCount >= 60 ? 'high' : 'medium', m.name, chrFirst);
    if (m.error) add('Damaged', m.name, m.error, 'medium', m.name, 0);
  }
  for (const s of result.xlm || []) {
    add('Excel 4.0', s.name, `${s.kind} (${s.state})`, s.state === 'very hidden' ? 'high' : 'medium', null, 0);
    for (const f of s.formulas || []) if (XLM_RISK.test(f.formula)) add('Excel 4.0', f.formula.match(XLM_RISK)[0].replace(/\s*\($/, ''), `Formula in ${s.name}!${f.cell}: ${preview(f.formula, 120)}`, /EXEC|CALL|REGISTER|URLDownload/i.test(f.formula) ? 'high' : 'medium', `${s.name}!`, 0);
  }
  for (const nm of result.xlmNames || []) add('Auto-run', nm.name, `Excel 4.0 auto-run name${nm.ref ? ` pointing to ${nm.ref}` : ''}`, 'high', null, 0);
  if (result.protected) add('Info', 'VBA password', 'The project has a VBA password. It only hides the code in Office; the source is still in the file and was scanned.', 'info', null, 0);

  const findings = [...map.values()].sort((a, b) => (SEV_RANK[b.severity] - SEV_RANK[a.severity]) || a.type.localeCompare(b.type) || a.keyword.localeCompare(b.keyword));
  const counts = { high: 0, medium: 0, low: 0, info: 0 };
  for (const f of findings) counts[f.severity]++;
  const auto = findings.some((f) => f.type === 'Auto-run' && f.severity === 'high');
  const risky = findings.some((f) => f.type !== 'Auto-run' && f.severity === 'high');
  let level = 'none';
  if (counts.high) level = 'high';
  else if (counts.medium) level = 'medium';
  else if (counts.low) level = 'low';
  let summary;
  const nMod = (result.modules || []).length;
  if (!nMod && !(result.xlm || []).length) summary = 'No VBA modules and no Excel 4.0 macro sheets: nothing to scan.';
  else if (auto && risky) summary = 'Code runs by itself when the file opens or closes, and it uses calls that can start programs, download files or run machine code. Check it before you enable macros.';
  else if (risky) summary = 'The macros use calls that can start programs, download files or change code. They do not run by themselves on open; check them before you run them.';
  else if (auto) summary = 'Code runs by itself when the file opens or closes. No calls that start programs or download files were found.';
  else if (level === 'medium') summary = 'Some calls deserve a look (COM objects, network, files or hidden text). No auto-run entry point was found.';
  else if (level === 'low') summary = 'Only low-risk findings, such as file access or environment reads.';
  else summary = 'No auto-run entry points and no risky calls were found.';
  return { file: fileName, findings, counts, level, autoRun: auto, summary, modules: nMod, lines: totalLines, xlmSheets: (result.xlm || []).length };
}

const esc = (s) => String(s).replace(/\|/g, '\\|').replace(/\r?\n/g, ' ');
/** Markdown report of a scan. */
function toMarkdown(rep, meta = {}) {
  const out = [];
  out.push(`# Macro security scan: ${meta.file || rep.file || 'file'}`);
  out.push('');
  if (meta.date) out.push(`Scanned ${meta.date} by VBA Viewer & Macro Extractor (Zovo). The macros were read as text; nothing was run.`);
  else out.push('The macros were read as text; nothing was run.');
  out.push('');
  out.push(`- Result: **${rep.level === 'none' ? 'no findings' : `${rep.level} risk`}**. ${rep.summary}`);
  out.push(`- Modules: ${rep.modules}, code lines: ${rep.lines}, Excel 4.0 macro sheets: ${rep.xlmSheets}`);
  out.push(`- Findings: ${rep.counts.high} high, ${rep.counts.medium} medium, ${rep.counts.low} low, ${rep.counts.info} info`);
  out.push('');
  if (rep.findings.length) {
    out.push('| Severity | Type | Keyword | Description | Count | Where |');
    out.push('|---|---|---|---|---:|---|');
    for (const f of rep.findings) {
      const where = f.places.filter((p) => p.module).slice(0, 5).map((p) => (p.line ? `${p.module}:${p.line}` : p.module)).join(', ') + (f.places.length > 5 ? ', ...' : '');
      out.push(`| ${f.severity} | ${esc(f.type)} | ${esc(f.keyword)} | ${esc(f.description)}${f.decoded ? ` (decoded: \`${esc(f.decoded).replace(/`/g, "'")}\`)` : ''} | ${f.count} | ${esc(where)} |`);
    }
    out.push('');
  }
  out.push('Findings are signs to check, not proof that the file is harmful or safe.');
  return `${out.join('\n')}\n`;
}

module.exports = { scan, toMarkdown, splitLine, SUSPICIOUS, AUTORUN, SEV_RANK };
