// Times loading a workbook (issue #5), and writes a load-and-save copy to compare part by part.
//   SaveBench load <xlsx> [runs] [copyPath]
using System;
using System.Diagnostics;
using System.IO;
using OfficeOpenXml;

static class LoadBench
{
    internal static void Run(string path, int runs, string copyPath)
    {
        for (int run = 0; run <= runs; run++)
        {
            var sw = Stopwatch.StartNew();
            using (var p = new ExcelPackage(new FileInfo(path)))
            {
                ExcelWorksheet ws = null; foreach (var w in p.Workbook.Worksheets) { ws = w; break; }
                var dim = ws.Dimension;          // forces the sheet to load
                long load = sw.ElapsedMilliseconds;
                if (run > 0) Console.WriteLine($"run {run}: load={load} ms  dimension={dim?.Address}");
                if (run == runs && copyPath != null) p.SaveAs(new FileInfo(copyPath));
            }
        }
    }
}
