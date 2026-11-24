// Save benchmark that mimics OraBooster's Excel mode (ClsExcel):
// cell-by-cell values, a date number format per column, a table, a bold header,
// AutoFitColumns on a sample, then SaveAs. Prints the time of each stage.
//   SaveBench [rows] [cols] [runs] [outDir]   benchmark
//   SaveBench keep <rows> <cols> <path>       write one workbook and keep it
//   SaveBench variety <path>                  write a workbook with many value kinds (see Variety.cs)
//   SaveBench issues <dir>                    write one workbook per fixed issue, to open in Excel (see Issues.cs)
//   SaveBench autofit | autofitsample | autofitrandom   AutoFitColumns: time and widths (see AutoFitBench.cs)
//   SaveBench charts <path>                   write one plain chart of every chart type, to compare two builds (see Charts.cs)
//   SB_FAST=1                                 benchmark with CompressionLevel.BestSpeed
using System;
using System.Diagnostics;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Table;

static class Program
{
    static int Main(string[] args)
    {
        if (args.Length == 2 && args[0] == "variety") { Variety.Write(args[1]); return 0; }
        if (args.Length == 2 && args[0] == "issues") { Issues.Write(args[1]); return 0; }
        if (args.Length == 2 && args[0] == "charts") { Charts.Write(args[1]); return 0; }
        if (args.Length >= 2 && args[0] == "load") { LoadBench.Run(args[1], args.Length > 2 ? int.Parse(args[2]) : 3, args.Length > 3 ? args[3] : null); return 0; }
        if (args.Length >= 1 && args[0] == "fill") { FillBench.Run(args.Length > 1 ? int.Parse(args[1]) : 50000, args.Length > 2 ? int.Parse(args[2]) : 15); return 0; }
        if (args.Length >= 1 && args[0] == "autofit") { AutoFitBench.Run(args.Length > 1 ? int.Parse(args[1]) : 50000, args.Length > 2 ? int.Parse(args[2]) : 15, args.Length > 3 ? int.Parse(args[3]) : 3); return 0; }
        if (args.Length == 2 && args[0] == "autofitsample") { AutoFitBench.WriteSample(args[1]); return 0; }
        if (args.Length == 3 && args[0] == "autofitrandom") { AutoFitBench.WriteRandom(args[1], int.Parse(args[2])); return 0; }
        if (args.Length == 4 && args[0] == "keep")
        {
            using (var p = new ExcelPackage())
            {
                Fill(p, int.Parse(args[1]), int.Parse(args[2]));
                Format(p, int.Parse(args[1]), int.Parse(args[2]));
                p.SaveAs(new FileInfo(args[3]));
            }
            return 0;
        }
        return Bench(args);
    }

    static void Fill(ExcelPackage p, int rows, int cols)
    {
        var ws = p.Workbook.Worksheets.Add("S");
        for (int c = 1; c <= cols; c++) ws.Cells[1, c].Value = "COL_" + c;
        var d0 = new DateTime(2026, 1, 1);
        for (int r = 2; r <= rows + 1; r++)
            for (int c = 1; c <= cols; c++)
            {
                object v;
                switch (c % 5)
                {
                    case 0: v = (double)(r * c); break;
                    case 1: v = "SQL_" + r.ToString("X8"); break;
                    case 2: v = d0.AddSeconds(r * 7 + c); break;
                    case 3: v = r % 7 == 0 ? null : "name " + (r % 1000); break;
                    default: v = r * 1.25 + c; break;
                }
                if (v != null) ws.Cells[r, c].Value = v;
            }
    }

    static void Format(ExcelPackage p, int rows, int cols)
    {
        var ws = p.Workbook.Worksheets["S"];
        for (int c = 2; c <= cols; c += 5) ws.Cells[2, c, rows + 1, c].Style.Numberformat.Format = "yyyy-mm-dd hh:mm:ss";
        var t = ws.Tables.Add(ws.Cells[1, 1, rows + 1, cols], "T1_1");
        t.TableStyle = TableStyles.Medium2;
        ws.Cells[1, 1, 1, cols].Style.Font.Bold = true;
    }

    static int Bench(string[] args)
    {
        int rows = args.Length > 0 ? int.Parse(args[0]) : 50000;
        int cols = args.Length > 1 ? int.Parse(args[1]) : 15;
        int runs = args.Length > 2 ? int.Parse(args[2]) : 3;
        string outDir = args.Length > 3 ? args[3] : Path.GetTempPath();
        Console.WriteLine($"EPPlus {typeof(ExcelPackage).Assembly.GetName().Name} {typeof(ExcelPackage).Assembly.GetName().Version}  rows={rows} cols={cols}");
        for (int run = 0; run <= runs; run++)
        {
            string path = Path.Combine(outDir, $"savebench_{run}.xlsx");
            var sw = Stopwatch.StartNew();
            long fill, fmt, fit, save;
            using (var p = new ExcelPackage())
            {
                if (Environment.GetEnvironmentVariable("SB_FAST") == "1") p.Compression = CompressionLevel.BestSpeed;
                Fill(p, rows, cols);
                fill = sw.ElapsedMilliseconds; sw.Restart();
                Format(p, rows, cols);
                fmt = sw.ElapsedMilliseconds; sw.Restart();
                p.Workbook.Worksheets["S"].Cells[1, 1, Math.Min(rows + 1, 1001), cols].AutoFitColumns(4, 80);
                fit = sw.ElapsedMilliseconds; sw.Restart();
                p.SaveAs(new FileInfo(path));
                save = sw.ElapsedMilliseconds;
            }
            long size = new FileInfo(path).Length;
            if (run > 0) Console.WriteLine($"run {run}: fill={fill} fmt+table={fmt} autofit={fit} save={save} total={fill + fmt + fit + save} ms  size={size}");
            File.Delete(path);
        }
        return 0;
    }
}
