// A workbook with one chart of every chart type the library can make, with nothing set but the series, a title and a style.
// Written with two builds of the library, the two files must be the same: a change in how charts are written shows here.
//   SaveBench charts <path>
using System;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

static class Charts
{
    internal static void Write(string path)
    {
        using (var p = new ExcelPackage())
        {
            var data = p.Workbook.Worksheets.Add("Data");
            string[] heads = { "Cat", "A", "B", "C", "D", "E" };
            for (int c = 0; c < heads.Length; c++) data.Cells[1, c + 1].Value = heads[c];
            var rnd = new Random(7);
            for (int r = 2; r <= 13; r++)
            {
                data.Cells[r, 1].Value = "c" + (r - 1);
                for (int c = 2; c <= 6; c++) data.Cells[r, c].Value = 10 + rnd.Next(90);
            }

            int made = 0, failed = 0;
            foreach (eChartType type in Enum.GetValues(typeof(eChartType)))
            {
                string name = type.ToString();
                var ws = p.Workbook.Worksheets.Add(name.Length > 31 ? name.Substring(0, 31) : name);
                try
                {
                    var chart = ws.Drawings.AddChart("Chart_" + name, type);
                    int series = name.StartsWith("StockVO") ? 5 : name.StartsWith("StockO") || name.StartsWith("StockV") ? 4 : 3;
                    for (int c = 0; c < series; c++)
                    {
                        string values = "'Data'!" + ExcelCellBase.GetAddress(2, c + 2, 13, c + 2);
                        ExcelChartSerie s = chart.Series is ExcelBubbleChartSeries
                            ? ((ExcelBubbleChartSeries)chart.Series).Add(values, "'Data'!A2:A13", "'Data'!F2:F13")
                            : chart.Series.Add(values, "'Data'!A2:A13");
                        s.HeaderAddress = data.Cells[1, c + 2];
                    }
                    chart.Title.Text = name;
                    chart.SetPosition(1, 0, 1, 0);
                    chart.SetSize(480, 300);
                    made++;
                }
                catch (Exception ex)
                {
                    // a chart type the library does not make: the sheet says so, and the two builds must agree on it
                    ws.Cells["A1"].Value = ex.GetType().Name + ": " + ex.Message;
                    failed++;
                }
            }
            p.SaveAs(new FileInfo(path));
            Console.WriteLine("charts: {0} made, {1} not made", made, failed);
        }
    }
}
