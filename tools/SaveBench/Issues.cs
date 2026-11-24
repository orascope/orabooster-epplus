// Writes one workbook per fixed issue, to open in Excel and check that it opens without a repair.
//   SaveBench issues <dir>
using System;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

static class Issues
{
    internal static void Write(string a_dir)
    {
        Directory.CreateDirectory(a_dir);

        // #2: delete a table on a workbook with several sheets, then add one
        using (var p = new ExcelPackage())
        {
            for (int s = 1; s <= 3; s++)
            {
                var ws = p.Workbook.Worksheets.Add("S" + s);
                ws.Cells["A1"].Value = "H"; ws.Cells["A2"].Value = s;
                ws.Tables.Add(ws.Cells["A1:A2"], "Table_S" + s);
            }
            p.Workbook.Worksheets["S1"].Tables.Delete("Table_S1");
            var ws1 = p.Workbook.Worksheets["S1"];
            ws1.Cells["C1"].Value = "H"; ws1.Cells["C2"].Value = 9;
            ws1.Tables.Add(ws1.Cells["C1:C2"], "Table_New");
            p.SaveAs(new FileInfo(Path.Combine(a_dir, "issue2_tables.xlsx")));
        }

        // #3: Column3D and the stock charts
        foreach (var type in new[] { eChartType.Column3D, eChartType.StockHLC, eChartType.StockOHLC })
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("S");
                string[] names = { "Open", "High", "Low", "Close" };
                int series = type == eChartType.StockOHLC ? 4 : 3;
                ws.Cells["A1"].Value = "Day";
                for (int c = 0; c < series; c++) ws.Cells[1, c + 2].Value = names[names.Length - series + c];
                var rnd = new Random(1);
                for (int r = 2; r <= 21; r++)
                {
                    ws.Cells[r, 1].Value = new DateTime(2026, 1, 1).AddDays(r);
                    ws.Cells[r, 1].Style.Numberformat.Format = "yyyy-mm-dd";
                    double open = 100 + rnd.Next(20), close = 100 + rnd.Next(20);
                    double high = Math.Max(open, close) + rnd.Next(5), low = Math.Min(open, close) - rnd.Next(5);
                    double[] v = series == 4 ? new[] { open, high, low, close } : new[] { high, low, close };
                    for (int c = 0; c < series; c++) ws.Cells[r, c + 2].Value = v[c];
                }
                var chart = ws.Drawings.AddChart("Chart_" + type, type);
                for (int c = 0; c < series; c++)
                {
                    var s = chart.Series.Add(ExcelCellBase.GetAddress(2, c + 2, 21, c + 2), "A2:A21");
                    s.HeaderAddress = ws.Cells[1, c + 2];
                }
                chart.Title.Text = type.ToString();
                chart.SetPosition(1, 0, 6, 0);
                chart.SetSize(600, 360);
                p.SaveAs(new FileInfo(Path.Combine(a_dir, "issue3_" + type + ".xlsx")));
            }
        }

        // #9: ignored errors, on a sheet that also has the parts written after them (drawing, comments, table).
        // Columns A and D fail an error check and are ignored (no green triangle); columns B and E fail the same checks and are not.
        using (var p = new ExcelPackage())
        {
            var ws = p.Workbook.Worksheets.Add("S");
            ws.Cells["A1"].Value = "Id (ignored)"; ws.Cells["B1"].Value = "Id"; ws.Cells["C1"].Value = "N";
            ws.Cells["D1"].Value = "Error (ignored)"; ws.Cells["E1"].Value = "Error";
            for (int r = 2; r <= 21; r++)
            {
                ws.Cells[r, 1].Value = "00" + r; ws.Cells[r, 2].Value = "00" + r; ws.Cells[r, 3].Value = r;
                ws.Cells[r, 4].Formula = "1/0"; ws.Cells[r, 5].Formula = "1/0";
            }
            ws.Tables.Add(ws.Cells["A1:E21"], "Table_Ids");
            ws.IgnoredErrors.Add("A2:A21", eErrorChecks.NumberStoredAsText);
            ws.IgnoredErrors.Add("D2:D21", eErrorChecks.EvalError);
            var chart = ws.Drawings.AddChart("Chart_N", eChartType.ColumnClustered);
            chart.Series.Add("C2:C21", "A2:A21");
            chart.SetPosition(1, 0, 7, 0);
            ws.Cells["A1"].AddComment("Text ids: the number-stored-as-text check is ignored for A2:A21", "SaveBench");
            ws.Cells["A1:E21"].AutoFitColumns();
            p.SaveAs(new FileInfo(Path.Combine(a_dir, "issue9_ignored_errors.xlsx")));
        }
    }
}
