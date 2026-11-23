# orabooster-epplus

A fork of **EPPlusFree 4.5.3.8** (EPPlus 4.5, LGPL) used by [OraBooster](https://www.orascope.net), with a faster save path.

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
<PackageReference Include="OraBooster.EPPlus" Version="4.5.3.8-ora.1" />
```

The package is not on nuget.org. Build it (below) or take the `.nupkg` from the [releases](https://github.com/orascope/orabooster-epplus/releases), and add the folder as a package source.

## What changed

Full list with file names: [CHANGES.md](CHANGES.md). Every change is also marked in the code with `orabooster-epplus:`.

| Version | Change |
|---|---|
| 4.5.3.8-ora.1 | Saves the package with `System.IO.Compression` (native zlib) instead of the managed DotNetZip deflater |
| | Fewer allocations when worksheet cells are written; `TypeCompat.IsPrimitive` without reflection |
| | `AutoFitColumns` disposes its GDI+ objects (they were left to the finalizer, one `StringFormat` per measured cell) |
| | Assembly and package `OraBooster.EPPlus`, `netstandard2.1` only, not strong-name signed, `System.Drawing.Common` 4.7.2 |

**Output.** For the same workbook, every part of the saved package (sheets, styles, shared strings, tables, drawings, properties) is byte-for-byte the same as EPPlusFree 4.5.3.8. Only the zip compression differs; files are a few percent larger.
Reading packages is unchanged (still DotNetZip).

**Speed.** `tools/SaveBench`, 50,000 rows x 15 columns set cell by cell, a table and AutoFit on 1,000 rows, the same PC:

| | EPPlusFree 4.5.3.8 | 4.5.3.8-ora.1 |
|---|---|---|
| `SaveAs` | 1.29-1.51 s | 0.78-1.02 s |

Most of the remaining save time is zlib. `ExcelPackage.Compression = CompressionLevel.BestSpeed` saves the same workbook in about 0.40-0.45 s with a file about 1.8x larger.

## Build

```
dotnet pack EPPlus/EPPlus.MultiTarget.csproj -c Release
```

Produces `EPPlus/bin/Release/OraBooster.EPPlus.<version>.nupkg` (0 errors; the CS1591 doc-comment warnings come from the original code).

Tools:
- `tools/SaveBench`: `SaveBench [rows] [cols] [runs]` times each stage of a save. `SaveBench variety <path>` and `SaveBench keep <rows> <cols> <path>` write sample workbooks. Build it with `-p:UsePackage=true` to run the same code against the EPPlusFree 4.5.3.8 package.
- `tools/SaveBench/compare-xlsx.ps1 -A a.xlsx -B b.xlsx`: compares two workbooks part by part (entry names and uncompressed bytes).

## Source of the OraBooster binaries

Each `OraBooster.EPPlus.dll` shipped with OraBooster is built from a tag of this repository, named after the package version (for example [`v4.5.3.8-ora.1`](https://github.com/orascope/orabooster-epplus/tree/v4.5.3.8-ora.1)).
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
