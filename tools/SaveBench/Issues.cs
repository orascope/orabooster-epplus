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

        // #11: the stock charts with volume. The series are added one after the other to the chart from AddChart,
        // volume first, as a caller does that knows nothing about the two parts of the chart.
        foreach (var type in new[] { eChartType.StockVHLC, eChartType.StockVOHLC })
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("S");
                string[] names = type == eChartType.StockVOHLC
                    ? new[] { "Volume", "Open", "High", "Low", "Close" }
                    : new[] { "Volume", "High", "Low", "Close" };
                ws.Cells["A1"].Value = "Day";
                for (int c = 0; c < names.Length; c++) ws.Cells[1, c + 2].Value = names[c];
                var rnd = new Random(1);
                for (int r = 2; r <= 21; r++)
                {
                    ws.Cells[r, 1].Value = new DateTime(2026, 1, 1).AddDays(r);
                    ws.Cells[r, 1].Style.Numberformat.Format = "yyyy-mm-dd";
                    double open = 100 + rnd.Next(20), close = 100 + rnd.Next(20);
                    double high = Math.Max(open, close) + rnd.Next(5), low = Math.Min(open, close) - rnd.Next(5);
                    double[] v = type == eChartType.StockVOHLC ? new[] { open, high, low, close } : new[] { high, low, close };
                    ws.Cells[r, 2].Value = 1000 + rnd.Next(9000);
                    for (int c = 0; c < v.Length; c++) ws.Cells[r, c + 3].Value = v[c];
                }
                var chart = ws.Drawings.AddChart("Chart_" + type, type);
                for (int c = 0; c < names.Length; c++)
                {
                    var s = chart.Series.Add(ExcelCellBase.GetAddress(2, c + 2, 21, c + 2), "A2:A21");
                    s.HeaderAddress = ws.Cells[1, c + 2];
                }
                chart.Title.Text = type.ToString();
                chart.SetPosition(1, 0, 7, 0);
                chart.SetSize(600, 360);

                // the same chart on a chart sheet
                var sheetChart = p.Workbook.Worksheets.AddChart("ChartSheet", type).Chart;
                for (int c = 0; c < names.Length; c++)
                {
                    var s = sheetChart.Series.Add(ws.Cells[2, c + 2, 21, c + 2], ws.Cells["A2:A21"]);
                    s.HeaderAddress = ws.Cells[1, c + 2];
                }
                p.SaveAs(new FileInfo(Path.Combine(a_dir, "issue11_" + type + ".xlsx")));
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

        // #19: the format of series and of single points, and error bars, one chart type per sheet (the sheet has the name of the type).
        // In every chart: series 1 has the fill C00000 and the border or line 0070C0, its third point has the fill (or the marker
        // fill) 00B050, and it has error bars where the chart type has them. chart-format-in-excel.ps1 reads these back from Excel.
        using (var p = new ExcelPackage())
        {
            var red = System.Drawing.Color.FromArgb(0xC0, 0x00, 0x00);
            var blue = System.Drawing.Color.FromArgb(0x00, 0x70, 0xC0);
            var green = System.Drawing.Color.FromArgb(0x00, 0xB0, 0x50);
            var types = new[]
            {
                eChartType.ColumnClustered, eChartType.BarStacked, eChartType.ColumnClustered3D, eChartType.Line, eChartType.LineMarkers,
                eChartType.Pie, eChartType.Doughnut, eChartType.Area, eChartType.XYScatter, eChartType.XYScatterLines, eChartType.RadarMarkers,
                eChartType.Bubble, eChartType.StockHLC, eChartType.StockVHLC
            };
            foreach (var type in types)
            {
                var ws = p.Workbook.Worksheets.Add(type.ToString());
                ws.Cells["A1"].Value = "Cat"; ws.Cells["B1"].Value = "V1"; ws.Cells["C1"].Value = "V2"; ws.Cells["D1"].Value = "V3"; ws.Cells["E1"].Value = "V4";
                for (int r = 2; r <= 7; r++)
                {
                    ws.Cells[r, 1].Value = r - 1; ws.Cells[r, 2].Value = r * 2 + 10; ws.Cells[r, 3].Value = 9 - r; ws.Cells[r, 4].Value = r + 3; ws.Cells[r, 5].Value = r + 1;
                }
                var chart = ws.Drawings.AddChart("Chart_" + type, type);
                bool stock = type == eChartType.StockHLC || type == eChartType.StockVHLC;
                int series = type == eChartType.StockVHLC ? 4 : type == eChartType.StockHLC ? 3 : 2;
                ExcelChartSerie first = null;
                for (int c = 0; c < series; c++)
                {
                    var s = type == eChartType.Bubble
                        ? ((ExcelBubbleChartSeries)chart.Series).Add(ExcelCellBase.GetAddress(2, c + 2, 7, c + 2), "A2:A7", "E2:E7")
                        : chart.Series.Add(ExcelCellBase.GetAddress(2, c + 2, 7, c + 2), "A2:A7");
                    s.HeaderAddress = ws.Cells[1, c + 2];
                    if (first == null) first = s;
                }
                bool markers = type == eChartType.Line || type == eChartType.LineMarkers || type == eChartType.XYScatter || type == eChartType.XYScatterLines || type == eChartType.RadarMarkers;
                first.Border.Fill.Color = blue;
                first.Border.Width = 2;
                if (!markers) first.Fill.Color = red;
                if (!stock)
                {
                    var point = first.DataPoints.Add(2);
                    if (markers) { point.MarkerFill.Color = green; point.MarkerBorder.Fill.Color = green; }
                    else point.Fill.Color = green;
                    first.DataPoints.Add(4).Border.Fill.Color = green;   // a second point, added after the third, with a border only
                    first.DataPoints.Add(0).Border.Width = 3;            // and one before them
                }
                bool bars = type == eChartType.ColumnClustered || type == eChartType.BarStacked || type == eChartType.Line || type == eChartType.LineMarkers ||
                            type == eChartType.Area || type == eChartType.XYScatter || type == eChartType.XYScatterLines || type == eChartType.Bubble;
                if (bars)
                {
                    var e = first.AddErrorBars(eErrorBarType.Both, eErrorValueType.FixedValue, 1.5);
                    e.Border.Fill.Color = blue;
                    if (type == eChartType.XYScatter || type == eChartType.Bubble) first.AddErrorBars(eErrorBarType.Plus, eErrorValueType.Percentage, 10, true);
                    chart.Series[1].AddErrorBars(eErrorBarType.Minus, eErrorValueType.StandardDeviation, 2);
                }
                // the second series of the scatter chart without lines gets a line through the property of the scatter series
                if (type == eChartType.XYScatter) ((ExcelScatterChartSerie)chart.Series[1]).LineColor = red;
                // the things a series could have before, to see that they still end up in a file Excel opens
                if (type == eChartType.ColumnClustered || type == eChartType.Line) first.TrendLines.Add(eTrendLine.Linear);
                if (type == eChartType.ColumnClustered) ((ExcelBarChartSerie)first).DataLabel.ShowValue = true;
                if (type == eChartType.Line) ((ExcelLineChartSerie)first).DataLabel.ShowValue = true;
                if (type == eChartType.Pie) { ((ExcelPieChartSerie)first).DataLabel.ShowPercent = true; ((ExcelPieChartSerie)first).Explosion = 10; }
                if (type == eChartType.XYScatterLines) ((ExcelScatterChartSerie)first).DataLabel.ShowValue = true;
                if (type == eChartType.LineMarkers)
                {
                    var line = (ExcelLineChartSerie)first;
                    line.LineColor = blue; line.Marker = eMarkerStyle.Circle; line.MarkerSize = 9; line.MarkerLineColor = red; line.Smooth = true;
                }
                if (type == eChartType.XYScatterLines)
                {
                    var scatter = (ExcelScatterChartSerie)first;
                    scatter.Marker = eMarkerStyle.Diamond; scatter.MarkerSize = 8; scatter.MarkerColor = red; scatter.LineColor = blue;
                }
                if (type == eChartType.RadarMarkers) ((ExcelRadarChartSerie)first).MarkerSize = 7;
                chart.SetPosition(1, 0, 6, 0);
                chart.SetSize(520, 320);
            }
            p.SaveAs(new FileInfo(Path.Combine(a_dir, "issue19_chart_format.xlsx")));
        }

        // #17: filters and sorts of an autofilter, one case per sheet. filters-in-excel.ps1 lists the hidden rows of each sheet,
        // lets Excel apply the filter again and lists them again: the two must be the same.
        using (var p = new ExcelPackage())
        {
            FilterCase(p, "values", false, f => f.ShowValues(0, "alpha", "gamma"));
            FilterCase(p, "values blank", false, f => f.ShowValues(3, "note 3", ""));
            FilterCase(p, "values number", false, f => f.ShowValues(1, "9.00", "15.00"));
            FilterCase(p, "values date", false, f => f.ShowValues(2, "2026-01-05", "2026-01-07"));
            FilterCase(p, "number and", false, f => f.ShowWhere(1, eFilterOperator.GreaterThanOrEqual, 6, true, eFilterOperator.LessThan, 27));
            FilterCase(p, "number equal", false, f => f.ShowWhere(1, eFilterOperator.Equal, 12));
            FilterCase(p, "number not equal", false, f => f.ShowWhere(5, eFilterOperator.NotEqual, 5));
            FilterCase(p, "number less mixed", false, f => f.ShowWhere(5, eFilterOperator.LessThan, 10));
            FilterCase(p, "text or blank", false, f => f.ShowWhere(3, eFilterOperator.Equal, "note 1*", false, eFilterOperator.Equal, ""));
            FilterCase(p, "text not equal", false, f => f.ShowWhere(4, eFilterOperator.NotEqual, "g0"));
            FilterCase(p, "text greater", false, f => f.ShowWhere(0, eFilterOperator.GreaterThan, "c"));
            FilterCase(p, "text wildcard mixed", false, f => f.ShowWhere(5, eFilterOperator.Equal, "*a*"));
            FilterCase(p, "text not wildcard mixed", false, f => f.ShowWhere(5, eFilterOperator.NotEqual, "5*"));
            FilterCase(p, "one character", false, f => f.ShowWhere(4, eFilterOperator.Equal, "g?"));
            FilterCase(p, "not blank", false, f => f.ShowWhere(3, eFilterOperator.NotEqual, ""));
            FilterCase(p, "blank", false, f => f.ShowWhere(3, eFilterOperator.Equal, null));
            FilterCase(p, "date from", false, f => f.ShowWhere(2, eFilterOperator.GreaterThanOrEqual, new DateTime(2026, 1, 6)));
            FilterCase(p, "date equal", false, f => f.ShowWhere(2, eFilterOperator.Equal, new DateTime(2026, 1, 5)));
            // Excel's own rules for "equal": the text the cell shows
            FilterCase(p, "equal shown text", false, f => f.ShowWhere(1, eFilterOperator.Equal, "12.00"));
            FilterCase(p, "equal two numbers", false, f => f.ShowWhere(1, eFilterOperator.Equal, 12, false, eFilterOperator.Equal, 15));
            FilterCase(p, "equal mixed", false, f => f.ShowWhere(5, eFilterOperator.Equal, 5, false, eFilterOperator.Equal, "abc"));
            FilterCase(p, "wildcard on numbers", false, f => f.ShowWhere(1, eFilterOperator.Equal, "1*", false, eFilterOperator.Equal, "2?.00"));
            FilterCase(p, "not equal number shown", false, f => f.ShowWhere(1, eFilterOperator.NotEqual, 12));
            FilterCase(p, "not equal two", false, f => f.ShowWhere(5, eFilterOperator.NotEqual, 5, true, eFilterOperator.NotEqual, "x"));
            FilterCase(p, "greater text mixed", false, f => f.ShowWhere(5, eFilterOperator.GreaterThan, "b"));
            FilterCase(p, "less or equal text", false, f => f.ShowWhere(0, eFilterOperator.LessThanOrEqual, "beta"));
            FilterCase(p, "tilde", false, f => f.ShowWhere(4, eFilterOperator.Equal, "g~?", false, eFilterOperator.Equal, "g1"));
            FilterCase(p, "several columns", false, f =>
            {
                f.ShowValues(0, "alpha", "gamma");
                f.ShowWhere(1, eFilterOperator.GreaterThanOrEqual, 6, true, eFilterOperator.LessThan, 27);
                f.ShowWhere(3, eFilterOperator.Equal, "note*", false, eFilterOperator.Equal, "");
                f.ShowWhere(4, eFilterOperator.NotEqual, "g0");
            });
            FilterCase(p, "sorted", false, f => f.Sort(1, true));
            FilterCase(p, "sorted and filtered", false, f => { f.ShowWhere(4, eFilterOperator.NotEqual, "g0"); f.Sort(0); });
            FilterCase(p, "table values", true, f => f.ShowValues(0, "alpha", "gamma"));
            FilterCase(p, "table several", true, f =>
            {
                f.ShowValues(3, "note 3", "note 6", "");
                f.ShowWhere(1, eFilterOperator.GreaterThan, 5);
            });
            FilterCase(p, "table sorted", true, f => { f.Sort(1, true); f.ShowWhere(4, eFilterOperator.Equal, "g1"); });
            FilterCase(p, "table cleared", true, f => { f.ShowValues(0, "alpha"); f.Clear(0); });
            p.SaveAs(new FileInfo(Path.Combine(a_dir, "issue17_autofilter.xlsx")));
        }
    }

    // A sheet with a header and ten rows: a name, a number shown with two decimals, a date, a note with empty cells,
    // a group, and a column that mixes numbers and texts. The autofilter is on a range or on a table.
    static void FilterCase(ExcelPackage p, string name, bool table, Action<ExcelAutoFilter> set)
    {
        var ws = p.Workbook.Worksheets.Add(name);
        string[] heads = { "Name", "Qty", "When", "Note", "Grp", "Mixed" };
        string[] names = { "alpha", "beta", "gamma", "alpha", "delta", "beta", "Alpha", "epsilon", "gamma", "zeta" };
        object[] mixed = { 5, "5a", 12, "abc", 7, "x", null, 5, "banana", 10.5 };
        for (int c = 0; c < heads.Length; c++) ws.Cells[1, c + 1].Value = heads[c];
        for (int i = 1; i <= 10; i++)
        {
            ws.Cells[i + 1, 1].Value = names[i - 1];
            ws.Cells[i + 1, 2].Value = i * 3;
            ws.Cells[i + 1, 3].Value = new DateTime(2026, 1, i);
            if (i % 4 != 0) ws.Cells[i + 1, 4].Value = "note " + i;
            ws.Cells[i + 1, 5].Value = "g" + (i % 3);
            if (mixed[i - 1] != null) ws.Cells[i + 1, 6].Value = mixed[i - 1];
        }
        ws.Cells["B2:B11"].Style.Numberformat.Format = "0.00";
        ws.Cells["C2:C11"].Style.Numberformat.Format = "yyyy-mm-dd";
        if (table)
        {
            var t = ws.Tables.Add(ws.Cells["A1:F11"], "Table_" + p.Workbook.Worksheets.Count);
            set(t.AutoFilter);
        }
        else
        {
            ws.Cells["A1:F11"].AutoFilter = true;
            set(ws.AutoFilter);
        }
    }
}
