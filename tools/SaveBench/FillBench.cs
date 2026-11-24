// Times three ways to put values into a worksheet (issue #4).
//   SaveBench fill [rows] [cols]
using System;
using System.Collections.Generic;
using System.Diagnostics;
using OfficeOpenXml;

static class FillBench
{
    static object Value(int r, int c)
    {
        switch (c % 5)
        {
            case 0: return (double)(r * c);
            case 1: return "SQL_" + r.ToString("X8");
            case 2: return new DateTime(2026, 1, 1).AddSeconds(r * 7 + c);
            case 3: return r % 7 == 0 ? null : "name " + (r % 1000);
            default: return r * 1.25 + c;
        }
    }

    internal static void Run(int rows, int cols)
    {
        var data = new List<object[]>(rows);
        for (int r = 0; r < rows; r++) { var a = new object[cols]; for (int c = 0; c < cols; c++) a[c] = Value(r + 2, c + 1); data.Add(a); }
        for (int run = 0; run < 4; run++)
        {
            long t1, t2, t3;
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("S");
                var sw = Stopwatch.StartNew();
                for (int r = 0; r < rows; r++) for (int c = 0; c < cols; c++) { var v = data[r][c]; if (v != null) ws.Cells[r + 2, c + 1].Value = v; }
                t1 = sw.ElapsedMilliseconds;
            }
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("S");
                var sw = Stopwatch.StartNew();
                for (int r = 0; r < rows; r++) for (int c = 0; c < cols; c++) { var v = data[r][c]; if (v != null) ws.SetValue(r + 2, c + 1, v); }
                t2 = sw.ElapsedMilliseconds;
            }
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("S");
                var sw = Stopwatch.StartNew();
                ws.Cells[2, 1].LoadFromArrays(data);
                t3 = sw.ElapsedMilliseconds;
            }
            if (run > 0) Console.WriteLine($"run {run}: Cells[r,c].Value={t1} SetValue={t2} LoadFromArrays={t3} ms");
        }
    }
}
