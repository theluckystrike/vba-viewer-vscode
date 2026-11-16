'use strict';
// Builds samples/sample-macros.xlsm (shipped; made up for this extension, the code is harmless) and the made-up test
// fixtures in test/fixtures. Run: node scripts/make-samples.js
const fs = require('fs');
const path = require('path');
const B = require('../test/support/build');
const { encodeCp } = require('../src/core/export');

const ROOT = path.join(__dirname, '..');
const CRLF = (s) => s.replace(/\r?\n/g, '\r\n');
const doc = (name, guid, body) => CRLF(`Attribute VB_Name = "${name}"
Attribute VB_Base = "0{${guid}}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
${body}`);
const cls = (name, body) => CRLF(`Attribute VB_Name = "${name}"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
${body}`);

const SAMPLE = [
  { name: 'ThisWorkbook', type: 'document', source: doc('ThisWorkbook', '00020819-0000-0000-C000-000000000046', `Option Explicit

Private Sub Workbook_Open()
    ' Shows the report month in the status bar when the workbook opens.
    Application.StatusBar = "Budget report for " & Format(Date, "mmmm yyyy")
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.StatusBar = False
End Sub
`) },
  { name: 'Sheet1', type: 'document', source: doc('Sheet1', '00020820-0000-0000-C000-000000000046', `Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Recalculates the summary when an amount in column C changes.
    If Not Intersect(Target, Me.Range("C2:C200")) Is Nothing Then
        BudgetReport.BuildSummary
    End If
End Sub
`) },
  { name: 'BudgetReport', type: 'standard', source: CRLF(`Attribute VB_Name = "BudgetReport"
Option Explicit

' Budget report helpers for the Northwind Labs sample workbook.
' The names and amounts in this sample are made up.

Public Const HELP_URL As String = "https://example.com/budget/help"
Private Const GREETING_HEX As String = "48656C6C6F2C20776F726C6421"

Public Sub BuildSummary()
    Dim totals As Object
    Dim cell As Range
    Dim key As Variant
    Dim row As Long

    Set totals = CreateObject("Scripting.Dictionary")
    For Each cell In Sheet1.Range("B2:B200")
        If Len(cell.Value) > 0 Then
            totals(cell.Value) = totals(cell.Value) + cell.Offset(0, 1).Value
        End If
    Next cell

    row = 2
    For Each key In totals.Keys
        Sheet1.Cells(row, 6).Value = key
        Sheet1.Cells(row, 7).Value = totals(key)
        row = row + 1
    Next key
    Sheet1.Range("F1").Value = "Prepared by " & Environ("USERNAME")
End Sub

Public Function MonthTotal(ByVal monthNo As Integer) As Currency
    Dim cell As Range
    For Each cell In Sheet1.Range("A2:A200")
        If IsDate(cell.Value) Then
            If Month(cell.Value) = monthNo Then MonthTotal = MonthTotal + cell.Offset(0, 2).Value
        End If
    Next cell
End Function

Public Function VatAmount(ByVal net As Double, Optional ByVal rate As Double = 0.19) As Double
    VatAmount = Round(net * rate, 2)
End Function
`) },
  { name: 'Formatting', type: 'standard', source: CRLF(`Attribute VB_Name = "Formatting"
Option Explicit

Public Sub ApplyCurrencyFormat(ByVal target As Range)
    With target
        .NumberFormat = "#,##0.00 [$EUR]"
        .HorizontalAlignment = xlRight
        .Font.Color = &H404040
    End With
End Sub

Public Sub HighlightOverBudget(ByVal limit As Currency)
    Dim cell As Range
    For Each cell In Sheet1.Range("C2:C200")
        If cell.Value > limit Then cell.Interior.Color = RGB(255, 235, 230)
    Next cell
End Sub
`) },
  { name: 'LineItem', type: 'class', source: cls('LineItem', `Option Explicit

Private pCategory As String
Private pAmount As Currency

Public Property Get Category() As String
    Category = pCategory
End Property

Public Property Let Category(ByVal value As String)
    pCategory = Trim$(value)
End Property

Public Property Get Amount() As Currency
    Amount = pAmount
End Property

Public Property Let Amount(ByVal value As Currency)
    If value < 0 Then Err.Raise 5, "LineItem", "Amount cannot be negative"
    pAmount = value
End Property
`) },
];

const FORM = { name: 'EntryForm', type: 'form', frame: CRLF(`VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EntryForm
   Caption         =   "New budget line"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   3
End
`), source: CRLF(`Attribute VB_Name = "EntryForm"
Attribute VB_Base = "0{9B8D1D51-5E0B-4E44-9A1C-5D2E7A3C1F00}{4F0E2C31-1A2B-4C3D-8E9F-0A1B2C3D4E5F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub OkButton_Click()
    Sheet1.Range("B2").Value = Me.CategoryBox.Text
    Unload Me
End Sub
`) };

function main() {
  const out = (rel, buf) => { const f = path.join(ROOT, rel); fs.mkdirSync(path.dirname(f), { recursive: true }); fs.writeFileSync(f, buf); return f; };
  const vba = B.buildVbaProject({ name: 'BudgetMacros', modules: SAMPLE });
  out('samples/sample-macros.xlsm', B.buildXlsm({ vba, sheets: ['Sheet1'] }));
  // fixtures (made up): a form, a password, a second version for compare, code page 1251, a package with no VBA,
  // a Word-style compound file, an Excel 4.0 macro sheet, broken files
  out('test/fixtures/forms.xlsm', B.buildXlsm({ vba: B.buildVbaProject({ name: 'BudgetMacros', modules: [...SAMPLE, FORM] }) }));
  out('test/fixtures/password.bin', B.buildVbaProject({ name: 'Locked', password: true, modules: [{ name: 'Module1', type: 'standard', source: 'Attribute VB_Name = "Module1"\r\nSub Hidden()\r\n    MsgBox "still readable"\r\nEnd Sub\r\n' }] }));
  const v2 = SAMPLE.filter((m) => m.name !== 'Formatting').map((m) => (m.name === 'BudgetReport' ? { ...m, source: m.source.replace('Optional ByVal rate As Double = 0.19', 'Optional ByVal rate As Double = 0.2').replace('Public Function VatAmount', "' VAT changed in 2026\r\nPublic Function VatAmount") } : m));
  v2.push({ name: 'Archive', type: 'standard', source: 'Attribute VB_Name = "Archive"\r\nSub MoveOldRows()\r\nEnd Sub\r\n' });
  out('test/fixtures/sample-macros-v2.xlsm', B.buildXlsm({ vba: B.buildVbaProject({ name: 'BudgetMacros', modules: v2 }) }));
  const cyr = 'Attribute VB_Name = "Отчет"\r\n\' Сумма по месяцам\r\nSub Итог()\r\n    MsgBox "Готово"\r\nEnd Sub\r\n';
  out('test/fixtures/cp1251.bin', B.buildVbaProject({ name: 'Проект', codePage: 1251, encode: (s) => encodeCp(s, 1251), modules: [{ name: 'Отчет', type: 'standard', source: cyr }] }));
  out('test/fixtures/no-macros.xlsm', B.buildXlsm({ sheets: ['Sheet1'] }));
  out('test/fixtures/word-style.doc', B.writeCfb({ children: [{ name: 'WordDocument', data: Buffer.alloc(600, 0x11) }, { name: 'Macros', children: (() => { const inner = require('../src/core/cfb').readCfb(B.buildVbaProject({ name: 'Project', modules: [{ name: 'NewMacros', type: 'standard', source: 'Attribute VB_Name = "NewMacros"\r\nSub AutoOpen()\r\n    Selection.TypeText "Hello"\r\nEnd Sub\r\n' }] })); return toTree(inner, inner.root).children; })() }] }));
  out('test/fixtures/xlm-hidden.xlsm', B.buildXlsm({ sheets: ['Data'], macroSheets: [{ name: 'Macro1', state: 'veryHidden', formulas: [['A1', '=ALERT("Sample macro sheet")'], ['A2', '=HALT()']] }], autoOpen: 'Macro1!$A$1' }));
  const good = fs.readFileSync(path.join(ROOT, 'samples/sample-macros.xlsm'));
  out('test/fixtures/broken/cut-off.xlsm', good.subarray(0, Math.floor(good.length * 0.6)));
  out('test/fixtures/broken/not-office.xls', Buffer.from('Date,Category,Amount\r\n2026-11-01,Travel,120.50\r\n'));
  out('test/fixtures/broken/random.doc', Buffer.from(Array.from({ length: 3000 }, (_, i) => (i * 7919 + 13) % 251)));
  const bin = B.buildVbaProject({ name: 'Broken', modules: [{ name: 'Module1', type: 'standard', source: 'Attribute VB_Name = "Module1"\r\nSub A()\r\nEnd Sub\r\n' }] });
  const cut = Buffer.from(bin.subarray(0, 1536));
  out('test/fixtures/broken/cut-vbaProject.bin', cut);
  console.log('samples and fixtures written');
}

function toTree(cfb, e) {
  return { name: e.name, children: cfb.children(e).map((c) => (c.type === 1 ? toTree(cfb, c) : { name: c.name, data: cfb.read(c) })) };
}

if (require.main === module) main();
module.exports = { SAMPLE, FORM };
