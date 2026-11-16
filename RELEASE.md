# VBA Viewer & Macro Extractor 0.1.0 (zovo.zovo-vba-viewer)

New listing, staged only. NICHE6 row 6, built by lane NEWB13.

- vsix: out/NEWB13-dist/zovo.zovo-vba-viewer/zovo.zovo-vba-viewer-0.1.0.vsix
- sha256: 0b2765025b3e7079652e518ce1be3f83fec3325db08b5cb5e96eb87695779234 (350,715 bytes, 40 files)
- Product repo: factory/products/vba-viewer-vscode, local commit only (no remote).
- Unit tests: 251 of 251 pass with `npm test` (scripts/test.js; `node --test` in the repo runs 254 of 254). They cover the compound file and ZIP readers, the VBA decompression, the dir and PROJECT streams, code pages, the .ppt and .doc storages, Excel 4.0 macro sheets, the scan, export and compare cores, the extension and the 3 Pro tools against a vscode stub (file A runs, then file B opens the Pro panel, for each tool; the count survives a new extension host), and the license and rating tests copied by apply.sh.
- Real-app gate: VS Code 1.142.0 on port 9446 through buildslot, run r1 on this exact vsix: 87 PASS, 0 FAIL (out/NEWB13-data/vba/vba-gate-r1.json and .tsv, screenshots in out/NEWB13-data/vba/shots/vba-r1-*.png). Dev runs: dev1 (71 PASS, 16 FAIL, all wrong gate expectations) and dev2 (87 PASS, 0 FAIL).

## What it is

- Display name: VBA Viewer & Macro Extractor (28 characters). The phrase "vba" comes first, and the description starts with "VBA viewer".
- Keywords: vba, excel macro, xlsm, macro, office. Engines: ^1.75.0, so it installs in Cursor and Windsurf.
- Free core: a read-only viewer for the VBA macros inside .xlsm, .xlsb, .xls, .xltm, .xlam, .docm, .dotm, .doc, .pptm, .ppt and vbaProject.bin files. It shows a module list grouped like the VBA editor (document, standard, class and form modules) and the source of each module with line numbers and VBA colors. It has search across all modules, a details panel (type, stream, size, procedures, attributes), Copy, Open as text, and a project overview (references, code page, VBA password state). It also lists Excel 4.0 macro sheets with their visibility, and their formulas in OOXML workbooks. The compound file, ZIP and MS-OVBA readers are our own code; there are no third-party libraries. Nothing runs, nothing is written to the file, and there are no network calls (webview CSP default-src 'none', nonce scripts).
- Editor priority: "default" only for vbaProject.bin. All Office types use "option" plus "Open VBA Macros" in the Explorer and editor title menus, so the XLSX, DOCX and PPTX viewers stay the default.
- Zovo Pro tools (1 free try each, module 2026-11-15.tries1, listing tag "vba", tries key vbaViewer.proTrial, TRIAL_USES = 1 in src/pro.js, no triesMax in proinfo/vba.json):
  - Export all modules: writes .bas, .cls and .frm files plus a .vba-export.json manifest to a folder. A re-export rewrites only the changed modules and lists the ones that are gone. Output is UTF-8 or the project's code page.
  - Macro security scan: auto-run entry points, Shell, CreateObject, PowerShell, downloads, Windows API and memory calls, hex and Base64 strings (decoded), URLs, IP addresses, file names and Excel 4.0 macro sheets. It shows a severity table and saves a Markdown or JSON report.
  - Compare macros: two Office files compared module by module (changed, only in A, only in B, the same, with lines added and removed), and VS Code's side-by-side diff for each module.
- Nothing pops up on open. The Pro panel shows only after a click on a Pro tool with no tries left.

## Samples used

- Shipped: samples/sample-macros.xlsm (ours, made by scripts/make-samples.js: 5 modules, harmless code, made-up names).
- Own fixtures (test/fixtures, not in the vsix): sample-macros-v2.xlsm, forms.xlsm, xlm-hidden.xlsm, no-macros.xlsm, cp1251.bin, password.bin, word-style.doc and the broken files.
- Real third-party files (downloaded raw from GitHub into newb13-work/vba/dl, never opened in Office or run; only parsed). Full list with URLs in out/NEWB13-scripts/vba/samples.tsv:
  - Apache POI test-data (Apache-2.0): SimpleMacro.xls, SimpleMacro.xlsm, SquareMacro.xls, 45431.xlsm (code page 950), xlmmacro.xlsm, Simple.xlsb, SimpleMacro.doc, SimpleMacro.docm, 60158.docm, HeaderWithMacros.doc, SimpleMacro.ppt, SimpleMacro.pptm, WithMacros.ppt (code page 936), testPPT.pptm, and the published SimpleMacro.vba sources.
  - VBA-tools (MIT): VBA-Web Blank.xlsm, VBA-Web Example.xlsm (40 modules), VBA-Web Specs.xlsm (37 modules), VBA-JSON Specs.xlsm, VBA-Dictionary Specs.xlsm, and the published JsonConverter.bas, WebResponse.cls and IWebAuthenticator.cls.

## Gate coverage (r1)

- Nothing appears on start.
- The sample opens with the chip, the module groups and order, and the source text equal to the core reader's output. The checks cover the syntax colors, the module details, a click on a procedure, the Attributes tick box, search with a hit click, the project overview, the document module, Open as text (Visual Basic, read-only) and the toolbar row "1 free try".
- vbaProject.bin opens by default. A .docm without an association opens in the default editor, not ours, and the editor title button then opens the viewer.
- 15 real files open. Module1 matches the published POI SimpleMacro.vba in .xls, .xlsm, .doc and .docm. WebResponse and IWebAuthenticator match the published VBA-Web .cls files, and JsonConverter matches the published JsonConverter.bas. The checks also cover code pages 936 and 950, the .ppt embedded storage, module counts of 8, 10, 37 and 40, and the Excel 4.0 sheet.
- Own fixtures: a very hidden XLM sheet, a form module, Cyrillic code page 1251, a VBA password, and a Word-style .doc.
- Files without VBA show the empty state (3 files, plus a .pptm through the Explorer). 5 broken files each show a plain message.
- Explorer right-click on BUDGET.XLSM (upper case) shows Open VBA Macros and the 3 "Zovo Pro: ..." entries. A folder shows none.
- Pro, for each tool: the first use on file A runs and its output is checked.
  - Export: 10 files on disk, and JsonConverter.bas equals the published file.
  - Scan: the table, plus the .md and .json reports on disk.
  - Compare: the table and the diff tab.
  - The try is counted after the result. The next use on file B opens the Pro panel after the click, and nothing is written.
  - The info tab lists the tools with "no free tries left" and "1 free try left".
  - After Developer: Reload Window all 3 tools show "0 tries left" and the Pro panel still opens.
- End checks: no error toasts, no CSP violations, no URL opened, no webview network requests, no extension-host sockets, no webview console errors, and no process left.

## Defects found and fixed

- Product (small UX): in Compare macros, the "Choose a file..." item vanished as soon as the user typed a filter in the file pick. It now stays visible (alwaysShow).
- Gate: dev1 had 16 wrong expectations, none of them product defects. The gate read an empty source line as the one-space placeholder. The size is shown as "1.4 KB". The headings are upper case through CSS. The search hits fall in 2 modules, not 3. VBA-Web Specs has 37 modules (its PROJECT stream lists 37). Document modules export as .cls. The .docm files have no editor association in the gate. Two vbaProject.bin tabs were open at once, so the gate read the wrong viewer. The quick pick list is virtualized. Each was checked against the core reader and the raw PROJECT stream before the fix.

## Known limits

- Forms: the code and form properties show and export, but not the form controls, and there is no .frx export.
- Only the source code is read, not the compiled p-code. If a tampered file has different p-code, the viewer shows the source.
- Excel 4.0 formulas are listed for .xlsm, .xltm and .xlam only. In .xls and .xlsb only the sheet names and visibility are listed.
- Files encrypted with a password to open cannot be read (a plain message says so).
- Compare macros compares two files. It does not compare against git HEAD, and the row's 4th tool (workspace scan as CSV) was not built.
- The README screenshot comes from gate run dev2 (same viewer code as r1, which differs only by the screenshot file). The Explorer side bar is open in it.
