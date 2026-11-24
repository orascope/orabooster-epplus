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
- **#11** Stock charts with volume: `AddChart(name, eChartType.StockVHLC)` and `StockVOHLC` (they threw `NotSupportedException`, #3).
  In the file such a chart is two charts in one plot area, as Excel writes it: a clustered column chart with the volume on the primary axes and a `c:stockChart` (high-low lines, and up/down bars with an open series) on the secondary axes.
  - `AddChart` returns the column chart (`ChartType` is `ColumnClustered`); the stock chart is `chart.PlotArea.ChartTypes[1]` and uses the secondary axes.
  - Series are added in the order Excel uses: the first series added to the returned chart is the volume, the following ones (open,) high, low, close go to the stock chart (`ExcelChartSeries.AddSeries`). A caller that adds its series one after the other needs no special code. `chart.Series` holds the volume series only.
  - A pivot table as the source is refused with `NotSupportedException` before anything is added to the drawing (a pivot chart can not be combined with another chart type).
  - `ExcelChart.UpdateStockSeriesCache` also writes the caches of a stock chart that is combined with another chart type (Excel needs them to turn on the high-low lines).
  - The close marker of a high-low-close chart (a dash) is chosen by the position of the series in its chart, not by its `idx`; the two differ when chart types are combined.
  - Loading: a stock chart with up/down bars is reported as `StockOHLC` (it was always `StockHLC`), and a loaded stock chart that is combined with another chart type has its chart type (it had 0).
  - Checked in Excel 16 with `SaveBench issues` and `open-in-excel.ps1` (which now prints the chart type and the chart groups of every chart): the files open without a repair, on a worksheet and on a chart sheet `ChartType` is `xlStockVHLC` (90) / `xlStockVOHLC` (91), the first chart group has the one volume series on the primary axes, the second has the price series on the secondary axes with `HasHiLoLines` on and `HasUpDownBars` on for VOHLC. The same charts made by Excel itself give the same lines, and the exported pictures show the columns, the high-low lines and the up/down bars.
  - `variety`, `keep` and the issue 2, 3 and 9 samples of `SaveBench` are part by part the same as before.
  - Found on the way, not fixed here: a loaded chart lists its first chart type twice in `PlotArea.ChartTypes` (#12).
  - A number of series that does not fit the chart type does not break the file: all four stock chart types with 0 to 6 series (28 workbooks) open in Excel without a repair. Excel just draws what it can.
  - After a load the series are added to the chart they are added to: the routing of the series is set up only by `AddChart`. A copy of the worksheet has the whole chart.
  - Tests: 5 new, 1 removed (the one that expected the exception). Whole suite: 1,170 passed, 10 failed, 214 skipped (the same 10 as before).
- **#14** `ExcelRangeBase.AutoFitColumns` is 3 to 7 times faster and gives every column exactly the width it gave before.
  - The cells are read from the cell store by row and column. The loop used `foreach (var cell in this)`, which makes a range object and an address string for every cell, and `cell.Style.WrapText`, which makes a style object to read one flag.
  - What depends only on the style (font, indent, rotation, wrap, number format) is looked up once per style. The number format was searched in the list of formats for every cell.
  - A text that was measured with the same font is not measured again (up to 100,000 texts of at most 256 characters per font and call).
  - In a text of printable ASCII the digits count as one character, so a column of numbers or dates is measured a few times, not once per cell. This is used only for a font that passes a check: every digit alone, and before and after every printable ASCII character, measures the same as `0` in its place (about 3,800 measurements, made once 1,000 texts with digits were measured, so a small range does not pay for it). A font with digits of several widths (Georgia) fails the check and is measured as before.
  - Measured with `SaveBench autofit` on the development PC while it was busy with other work, the two builds run one after the other, three rounds: 50,000 rows x 15 columns 3.2-5.8 s -> 0.46-1.47 s; 1,000 rows 83-200 ms -> 24-119 ms; 10 rows 2 ms -> 1 ms.
  - Same widths: `SaveBench variety`, `keep`, `autofitsample`, the issue samples and twelve `SaveBench autofitrandom` workbooks (three sheets of 24 columns and 1,500 to 2,500 rows each, ten fonts, number and date formats, indents, rotated, wrapped and merged cells, a hidden column, tables and filters, a width limit, a range of several addresses), written by the build before and by this one, are part by part the same.
  - Not changed, measured for the record: against Excel's own autofit of the same cells (`SaveBench autofitsample`, then `autofit-in-excel.ps1`; Excel 16 at 125 % display scale) the widths are 6 % wider on average, and 4 of 18 columns are narrower, by up to 2.6 % (one character in 40). Excel measures with fractional glyph widths at the display scale of the PC; neither GDI+ (used here) nor GDI at 96, 120 or 144 dpi gives its numbers, so no model of it was added. A caller that must not cut a long text can widen the columns after the call (3 % covers the sample).
  - Found on the way, not fixed here: `AutoFitColumns` throws when a cell with a date format holds a number that is not a date (#15).
  - `tools/SaveBench`: `autofit`, `autofitsample`, `autofitrandom` (new `AutoFitBench.cs`) and `autofit-in-excel.ps1`.
  - Tests: 2 new. Whole suite: 1,172 passed, 10 failed, 214 skipped (the same 10 as before).
- **#17** Filters and sort of an autofilter: `ExcelWorksheet.AutoFilter` and `ExcelTable.AutoFilter` (new `ExcelAutoFilter.cs`: `ExcelAutoFilter`, `eFilterOperator`). Sorting a range was already there (`ExcelRangeBase.Sort`).
  - `ShowValues(column, texts)`: the rows whose cell shows one of the texts (`filters` / `filter`; an empty text is `blank="1"`). `ShowWhere(column, operator, value)` and the form with two conditions joined by and / or (`customFilters` / `customFilter`). `Clear`, `IsFiltered`. The column is zero based within the autofilter, as in `ExcelRangeBase.Sort`.
  - `Sort(column, descending)` sorts the rows below the buttons and records the sort so that the button shows the arrow: `sortState` inside `autoFilter` on a worksheet, after `autoFilter` in a table, as Excel writes them. `ClearSort` removes the record.
  - Excel does not apply a filter when it opens a file, so the rows that do not pass are hidden here: when the package is saved, if a filter was set or taken off through this class, or when `Apply` is called. A sheet where the class was not used is not touched.
  - A filter of a kind that is not handled (top 10, dynamic, colour, icon, date groups) is kept as it is and left to Excel: its column lets every row pass and no row is shown that was hidden.
  - The conditions follow what Excel 16 does, which is not always what one would guess: "equal" compares with the text the cell shows (a 12 shown as 12.00 is not equal to 12), "not equal" and the greater / less operators compare the number; a number cell does not match a text with wildcards; "not empty" is "not equal" to one space. For `ShowWhere(column, Equal, number or date)` the class writes "not less and not greater", which compares the number whatever the format is.
  - The text a cell shows depends on the culture of the thread (12.00 or 12,00), as it does in Excel; the texts given to `ShowValues` are compared with it.
  - Checked in Excel 16 with `SaveBench issues` and the new `tools/SaveBench/filters-in-excel.ps1`: for each of the 34 sheets of `issue17_autofilter.xlsx` (value lists, empty cells, numbers as they are shown, dates, one and two conditions with every operator, wildcards, a column of numbers and texts, sorts, tables) the rows hidden in the file are exactly the rows Excel hides when it applies the filter again; Excel sees the filters on the same columns and the same sort; the file opens without a repair.
  - `variety`, `keep`, the autofit samples and the issue 2, 3, 9 and 11 samples of `SaveBench` are part by part the same as before.
  - Not done: when rows or columns are inserted or deleted after `Sort`, the recorded sort is not moved (call `Sort` again or `ClearSort`).
  - Tests: 6 new. Whole suite: 1,178 passed, 10 failed, 214 skipped (the same 10 as before).

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