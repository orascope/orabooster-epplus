# Changes in orabooster-epplus

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