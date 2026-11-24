# Changes in orabooster-epplus

## 4.5.3.8-ora.3 (not released yet, OraScope Co.,Ltd.)

New features, taken one at a time from the candidate list in #8. They are written from ECMA-376 and the Microsoft Open Specifications and checked against Excel. A workbook that does not use them is saved exactly as by ora.2.

- **#9** Ignored errors: `ExcelWorksheet.IgnoredErrors` (new `ExcelIgnoredError.cs`: `ExcelIgnoredErrorCollection`, `ExcelIgnoredError`, `eErrorChecks`).
  Lists ranges for which Excel does not show given error checks (the green triangles), the `ignoredErrors` element of the worksheet part:
  `ws.IgnoredErrors.Add("A2:A100", eErrorChecks.NumberStoredAsText)`. The nine checks of the schema can be combined; an address can hold several ranges.
  - Entries of a loaded workbook are read; entries that are not changed are saved as they were.
  - `InsertRow`, `InsertColumn`, `DeleteRow` and `DeleteColumn` move the ranges; a range whose cells are all deleted is removed, and the element with the last one (the schema does not allow it empty).
  - Whole rows and columns are written with their first and last cell (`D1:D1048576`), as Excel does.
  - Checked in Excel 16: the sample of `SaveBench issues` (table, chart and comment on the same sheet) opens without a repair; `Range.Errors(xlNumberAsText).Ignore` and `Errors(xlEvaluateToError).Ignore` are true on the listed ranges and false on the same values next to them; Excel keeps the entries when it saves the file and adds its own in the same form. A copy with the element out of schema order does not open, so the check tells the difference.
  - `variety`, `keep` and the issue 2 and 3 samples of `SaveBench` are part by part the same as with ora.2.
  - The element is placed after the last element that the schema puts before it, not before the first later element the library knows: a loaded sheet whose later elements are not in the node order list (`controls`, `drawingHF`) still gets it at the right place.
  - An entry of a loaded file whose address cannot be read is left as it is and does not stop an insert or delete. `Add` takes a range object with several ranges and refuses one of another worksheet; an item cannot be left without a check (remove it instead).
  - `tools/SaveBench/open-in-excel.ps1` (new): opens every workbook of a folder in a hidden Excel and reports the ones that need a repair; used for the check above.
  - Tests: 9 new in `OraBoosterIssuesTests`. Whole suite: 1,166 passed, 10 failed, 214 skipped (the same 10 as before).

## 4.5.3.8-ora.2 (2026-11-24, OraScope Co.,Ltd.)

Fixes for the issues found in OraBooster. For workbooks without the fixed cases, every saved part is still byte-for-byte the same as EPPlusFree 4.5.3.8.

- **#2** `Table/ExcelTableCollection.Delete`: the next table id was decremented once per worksheet instead of once, so with several sheets a new table got id 0 or an id in use and Excel asked to repair the file. The next id now comes from the largest id left. The same rule as `ExcelWorkbook.ReadAllTables` (chart sheets skipped).
- **#3** Charts:
  - `Column3D` is a `c:bar3DChart` (it threw `NotImplementedException`).
  - `StockHLC` and `StockOHLC` are `c:stockChart` with high-low lines (and up/down bars for OHLC), no `varyColors`, series without their own line, the close of HLC as a dash.
    Excel turns on the high-low lines only when the series carry cached values when the file is opened, so `ExcelChart.UpdateStockSeriesCache` writes `numCache`/`strCache` from the worksheet when a stock chart is saved. Other chart types are saved as before.
  - `StockVHLC` and `StockVOHLC` (with volume) need a bar chart and a stock chart on two axis groups and are not supported. `AddChart` now throws `NotSupportedException` before anything is added to the drawing; `StockVHLC` used to fail half way and leave a broken drawing.
  - Checked in Excel: the files open without a repair, `ChartType` is `xl3DColumn` / `xlStockHLC` / `xlStockOHLC`, `HasHiLoLines` and `HasUpDownBars` are on, and the exported pictures show the bars and lines.
- **#4** `ExcelCellBase.GetColumnLetter` / `GetAddress`: built in one buffer instead of one string per letter (same text). `Cells[r, c].Value = v` is 5-10 % faster and saving uses them for every cell. The rest of the cost is the range object of the indexer: `ExcelWorksheet.SetValue(row, col, value)` puts the same values about 2.5x faster (`tools/SaveBench fill`).
- **#5** `Packaging/ZipPackage`: loads with `System.IO.Compression.ZipArchive` instead of DotNetZip's `ZipInputStream` (same rules for empty entries, separators and relations). 50,000 rows x 15 columns: 0.68-1.04 s -> 0.60-0.78 s; the rest is XML parsing. Load-and-save copies are part by part the same as with EPPlusFree.
- **#6** `EPPlusTest/EPPlusTest.Core.csproj` targets `net10.0-windows` with MSTest 3.6.4; two tests fixed for the newer compiler/MSTest; `TestBase` writes to `%TEMP%\EPPlusTest` instead of `C:\epplusTest`. New `OraBoosterIssuesTests` (5 tests, all fail on EPPlusFree 4.5.3.8).
  Result: 1,157 passed, 10 failed, 214 skipped; the same 10 fail on EPPlusFree 4.5.3.8 with the same test project. `WorkSheetTest.DateFunctionsWorkWithDifferentCultureDateFormats_GB` crashes the test host on both and is excluded.
## 4.5.3.8-ora.1 (2026-11-24, OraScope Co.,Ltd.)

Based on EPPlusFree 4.5.3.8 (https://github.com/rimland/EPPlus, commit c777bbe).

Faster saving. For the same workbook every part of the package is byte-for-byte the same as EPPlusFree 4.5.3.8;
only the zip compression differs (a few percent larger file).

- `Packaging/ZipOutputWriter.cs` (new): writes the package with `System.IO.Compression.ZipArchive` (native zlib).
  `ZipPackage`, `ZipPackagePart`, `ZipPackageRelationshipCollection`, `ExcelWorksheet.SaveHandler` and
  `ExcelWorkbook.SaveSharedStringHandler` use it instead of the managed DotNetZip `ZipOutputStream`.
  Reading packages still uses DotNetZip.
- `ExcelWorksheet.UpdateRowCellData` / `WriteRow` / `GetFormulaValue`: `Append` instead of `AppendFormat`,
  one dictionary lookup per shared string, no extra `ToString` for numbers, no copy of the 6 MB buffer into one string.
- `Compatibility/TypeCompat.IsPrimitive`: type tests instead of reflection (same result, also for null and enums).
- `ExcelRangeBase.AutoFitColumns`: disposes its `Bitmap`, `Graphics`, `Font` and `StringFormat` objects.
- Project: assembly and package renamed to `OraBooster.EPPlus`; targets `netstandard2.1` only;
  not strong-name signed; `System.Drawing.Common` 4.7.0 -> 4.7.2 (GHSA-rxg9-xrhp-64gj).

Measured with `tools/SaveBench` (50,000 rows x 15 columns, values set cell by cell, table, AutoFit on 1,000 rows):
`SaveAs` 1.29-1.51 s -> 0.78-1.02 s on the same PC.