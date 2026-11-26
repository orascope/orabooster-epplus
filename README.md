# orabooster-epplus

A fork of **EPPlusFree 4.5.3.8** (EPPlus 4.5, LGPL) used by [OraBooster](https://www.orascope.net), with a faster save path, fixes, and the features OraBooster needed that EPPlus 4.5 does not have.

It is not an official EPPlus release and is not supported by EPPlus Software AB or by the EPPlusFree author.
It is maintained only as far as OraBooster needs it.

## Origin

| | Repository | Version | License |
|---|---|---|---|
| Original | https://github.com/JanKallman/EPPlus (archived) | EPPlus 4.5.3.3, the last LGPL version | LGPL |
| Continued by | https://github.com/rimland/EPPlus | EPPlusFree 4.5.3.8 (commit `c777bbe`) | LGPL-3.0 |
| This fork | https://github.com/orascope/orabooster-epplus | `OraBooster.EPPlus` 4.5.3.8-ora.N | LGPL-3.0 |

EPPlus 5 and later (https://github.com/EPPlusSoftware/EPPlus) are under the Polyform Noncommercial license and are not used here.

## Requirements

- Target framework: `netstandard2.1`. Used by OraBooster on **.NET 10** (`net10.0-windows`); any .NET Core 3.0+ / .NET 5+ runtime can load it.
- Windows for the parts that use `System.Drawing` (`AutoFitColumns`, pictures). Reading and writing cells, styles, tables and charts do not need it.

## Use

The assembly is `OraBooster.EPPlus.dll`; the namespace is `OfficeOpenXml` as in EPPlus 4.5 and EPPlusFree, so code written for EPPlusFree 4.5.3.8 compiles unchanged.

```xml
<PackageReference Include="OraBooster.EPPlus" Version="4.5.3.8-ora.4" />
```

The package is not on nuget.org. Build it (below) or take the `.nupkg` from the [releases](https://github.com/orascope/orabooster-epplus/releases), and add the folder as a package source.

## What changed

Full list with file names: [CHANGES.md](CHANGES.md). Every change is also marked in the code with `orabooster-epplus:`.

| Version | Change |
|---|---|
| 4.5.3.8-ora.5 (not released yet) | The export of a range: `range.ToDataTable()`, `range.ToText(format)` and `range.SaveToText(file)` (CSV, the reverse of `LoadFromText`), `range.ToJson()`, `range.ToHtml()` (#32) |
| | Asynchronous input and output: `await package.SaveAsAsync(stream)`, `SaveAsAsync(file)`, `SaveAsync()`, `LoadAsync(stream)`, `LoadAsync(file)`. A stream that allows no synchronous writes (an ASP.NET Core response) can be saved to. The package itself is made in memory on the calling thread (#34) |
| | Fix: an encrypted workbook is loaded again on .NET 6 and later, also one that Excel encrypted (#35) |
| | Fix: `TableStyle` of a loaded table or pivot table has the style of the file; it was `Medium6` for every one (#39) |
| | Fix: a conditional format of a loaded workbook keeps its fill and its bold when its style is changed, and a double underline is written as one (#40) |
| 4.5.3.8-ora.4 | A chart series with no X range: `chart.Series.Add("B2:B20", null)` no longer throws, and `""` no longer writes a reference with no address (#27) |
| | A loaded `XYScatter` chart has its chart type (it was `XYScatterLinesNoMarkers`), so a series added to it has markers and no line (#29) |
| 4.5.3.8-ora.3 | Ignored errors: `ws.IgnoredErrors.Add("A2:A100", eErrorChecks.NumberStoredAsText)` hides the green error-check triangles on a range (#9) |
| | Stock charts with volume: `AddChart(name, eChartType.StockVHLC)` and `StockVOHLC`; add the series volume first, then (open,) high, low, close (#11) |
| | `AutoFitColumns` 3 to 7 times faster, with the same widths (#14) |
| | Filters and sort of an autofilter: `ws.AutoFilter.ShowValues(0, "alpha")`, `table.AutoFilter.ShowWhere(1, eFilterOperator.GreaterThan, 5)`, `.Sort(1, true)`; the rows that do not pass are hidden when the package is saved (#17) |
| | Chart series: `serie.Fill` / `serie.Border` no longer make a file Excel does not open on scatter and stock charts, and are written in the order of the schema; `serie.DataPoints.Add(2).Fill.Color` for one point; `serie.AddErrorBars(...)` (#19) |
| | The chart types of Excel 2016: `ws.Drawings.AddExtendedChart("name", eChartExType.Waterfall)` and `Treemap`, `Sunburst`, `Histogram`, `Pareto`, `BoxWhisker`, `Funnel`, then `chart.Series.Add("B2:B20", "A2:A20")` (#21) |
| | Fixes: `FormulaR1C1` in cultures with another minus sign (#23); loaded charts list each chart type once and series get free numbers (#12); a number that is not a date in a date format no longer stops `AutoFitColumns` (#15); the whole test suite passes |
| 4.5.3.8-ora.2 | Deleting a table no longer gives the next table id 0 or an id in use (#2) |
| | `Column3D`, `StockHLC` and `StockOHLC` charts; the volume stock charts throw `NotSupportedException` before anything is added (#3) |
| | Cell addresses without extra strings (#4); loading with `System.IO.Compression` (#5); EPPlusTest runs on .NET 10 (#6) |
| 4.5.3.8-ora.1 | Saves the package with `System.IO.Compression` (native zlib) instead of the managed DotNetZip deflater |
| | Fewer allocations when worksheet cells are written; `TypeCompat.IsPrimitive` without reflection |
| | `AutoFitColumns` disposes its GDI+ objects (they were left to the finalizer, one `StringFormat` per measured cell) |
| | Assembly and package `OraBooster.EPPlus`, `netstandard2.1` only, not strong-name signed, `System.Drawing.Common` 4.7.2 |

**Output.** For the same workbook (outside the cases fixed in ora.2), every part of the saved package (sheets, styles, shared strings, tables, drawings, properties) is byte-for-byte the same as EPPlusFree 4.5.3.8. Only the zip compression differs; files are a few percent larger.
Since ora.2 packages are also read with `System.IO.Compression`; a load-and-save copy is part by part the same as with EPPlusFree.
ora.3 saves a workbook that uses none of its new features and none of its fixed cases exactly as ora.2; [CHANGES.md](CHANGES.md) lists what a caller can see differently.
ora.4 saves a workbook exactly as ora.3, except for a chart series added with `""` as its X range.

**Speed.** `tools/SaveBench`, 50,000 rows x 15 columns set cell by cell, a table and AutoFit on 1,000 rows, the same PC:

| | EPPlusFree 4.5.3.8 | 4.5.3.8-ora.1 |
|---|---|---|
| `SaveAs` | 1.29-1.51 s | 0.78-1.02 s |

Loading 50,000 rows x 15 columns: 0.68-1.04 s -> 0.60-0.78 s (ora.2).

**Many values.** `ExcelWorksheet.SetValue(row, col, value)` puts values about 2.5x faster than `Cells[row, col].Value = value` (the indexer sets up a range object for every call): 100,000 x 15 values in about 80 ms instead of 200 ms (`tools/SaveBench fill`).

**AutoFit.** `AutoFitColumns` over 50,000 rows x 15 columns: 3.2-5.8 s -> 0.46-1.47 s (ora.3, `tools/SaveBench autofit`, both builds run one after the other on a busy PC), with the same widths. Against Excel's own autofit the widths are 6 % wider on average and up to 2.6 % narrower for some long texts (`SaveBench autofitsample`, then `autofit-in-excel.ps1`); widen the columns after the call if a long text must not be cut.

Most of the remaining save time is zlib. `ExcelPackage.Compression = CompressionLevel.BestSpeed` saves the same workbook in about 0.40-0.45 s with a file about 1.8x larger.

## Build

```
dotnet pack EPPlus/EPPlus.MultiTarget.csproj -c Release
```

Produces `EPPlus/bin/Release/OraBooster.EPPlus.<version>.nupkg` (0 errors; the CS1591 doc-comment warnings come from the original code).

Tests (EPPlusTest, MSTest, `net10.0-windows`):

```
dotnet test EPPlusTest/EPPlusTest.Core.csproj -c Release
```

1,213 passed, 0 failed, 214 skipped (the skipped ones are marked `[Ignore]` in EPPlus 4.5; most need files that are not in the repository). `OraBoosterIssuesTests` covers the fixes and the new features in this fork. Some old tests still write files to `C:\temp`.
Until ora.2 ten tests failed and one stopped the test host and was left out of the run; see #23 in [CHANGES.md](CHANGES.md).

Tools:
- `tools/SaveBench`: `SaveBench [rows] [cols] [runs]` times each stage of a save. `SaveBench variety <path>`, `SaveBench keep <rows> <cols> <path>` and `SaveBench issues <dir>` write sample workbooks; `SaveBench fill` and `SaveBench load <xlsx>` time putting values and loading. Build it with `-p:UsePackage=true` to run the same code against the EPPlusFree 4.5.3.8 package.
- `tools/SaveBench/compare-xlsx.ps1 -A a.xlsx -B b.xlsx`: compares two workbooks part by part (entry names and uncompressed bytes).
- `tools/SaveBench/open-in-excel.ps1 -Dir <folder>`: opens every workbook of a folder in a hidden Excel and reports the ones that need a repair (for the folder of `SaveBench issues`).
- `SaveBench autofit [rows] [cols] [runs]` times `AutoFitColumns`; `SaveBench autofitrandom <path> <seed>` writes a workbook of random cells to compare two builds with `compare-xlsx.ps1`; `SaveBench autofitsample <path>` and `tools/SaveBench/autofit-in-excel.ps1 -Path <xlsx>` compare the widths with Excel's own autofit.
- `tools/SaveBench/filters-in-excel.ps1 -Path <xlsx>`: for every sheet with an autofilter, the rows hidden in the file against the rows Excel hides when it applies the filter again (for `issue17_autofilter.xlsx` of `SaveBench issues`).
- `SaveBench charts <path>` writes one plain chart of every chart type, to compare two builds with `compare-xlsx.ps1`; `tools/SaveBench/chart-format-in-excel.ps1 -Path <xlsx>` reads the colours of series and points and the error bars back from Excel (for `issue19_chart_format.xlsx` of `SaveBench issues`).

## Source of the OraBooster binaries

Each `OraBooster.EPPlus.dll` shipped with OraBooster is built from a tag of this repository, named after the package version (for example [`v4.5.3.8-ora.2`](https://github.com/orascope/orabooster-epplus/tree/v4.5.3.8-ora.2)).
The dll is a separate file next to `OraBooster.exe`, so it can be replaced with a modified build of this source.

## License

GNU Lesser General Public License v3.0: [COPYING.LESSER](COPYING.LESSER), with the GNU General Public License v3.0 it refers to: [COPYING](COPYING).
The copyright notices of the original authors (Jan Källman, EPPlus Software AB and contributors, Rimland) are kept in the source files;
changes in this fork are (c) 2026 OraScope Co.,Ltd. under the same license.

## About EPPlus 4.5

EPPlus reads and writes Excel 2007+ files (Office Open XML, xlsx) without Excel or interop: cell ranges, styling, data validation,
conditional formatting, charts, pictures, shapes, comments, tables, pivot tables, protection, encryption, VBA and formula calculation.

The README of EPPlusFree, with the history of the project, is kept at
https://github.com/orascope/orabooster-epplus/blob/c777bbe/README.md (the original repository: https://github.com/JanKallman/EPPlus).
