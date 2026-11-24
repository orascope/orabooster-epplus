// Writes one workbook with many value kinds, styles, formulas, a table and a chart.
// Used to check that a change to the save path keeps every package part the same.
using System;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

static class Variety
{
    enum Color3 { Red, Green, Blue }

    internal static void Write(string a_path)
    {
        using (var p = new ExcelPackage())
        {
            p.Workbook.Properties.Title = "variety";
            p.Workbook.Properties.Author = "SaveBench";
            p.Workbook.Properties.Created = new DateTime(2026, 1, 2, 3, 4, 5);
            var ws = p.Workbook.Worksheets.Add("Values");
            object[] vals =
            {
                "plain", " lead space", "two  spaces", "tab\there", "line\nbreak", "a&b<c>\"d'", "한글 값", "",
                true, false, 1, -2L, (short)3, (byte)4, (sbyte)-5, (ushort)6, 7u, 8ul, 9.5f, 10.25, -0.0, 1e-300, 1.7976931348623157e308,
                double.NaN, double.PositiveInfinity, double.NegativeInfinity, 12345678901234567890m, 0.1m,
                new DateTime(2026, 11, 24, 13, 45, 6), new DateTime(1900, 1, 1), TimeSpan.FromHours(26.5), 'x', Color3.Blue,
                new int[] { 7, 8 }, Guid.Empty
            };
            for (int i = 0; i < vals.Length; i++)
            {
                ws.Cells[i + 1, 1].Value = vals[i] == null ? "null" : vals[i].GetType().Name;
                ws.Cells[i + 1, 2].Value = vals[i];
            }
            ws.Cells[1, 3].Formula = "SUM(B9:B20)";
            ws.Cells[2, 3, 10, 3].Formula = "B2&\"x\"";
            ws.Cells[11, 3].Formula = "\"a&b\"";
            ws.Cells[12, 3].Value = 5; ws.Cells[12, 3].Formula = "2+3";
            ws.Cells[13, 3].Formula = "1/0";
            ws.Cells[1, 4, 20, 4].Style.Numberformat.Format = "#,##0.00";
            ws.Cells[1, 5].Style.Font.Bold = true;              // styled empty cell
            ws.Cells[2, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[2, 5].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
            ws.Cells[3, 5].RichText.Add("rich ").Bold = true;
            ws.Cells[3, 5].RichText.Add("text");
            ws.Row(4).Height = 30; ws.Row(5).Hidden = true;
            ws.Column(6).Width = 20;
            ws.Cells[30, 8].Value = "far away";

            var t = p.Workbook.Worksheets.Add("Table");
            t.Cells[1, 1].Value = "K"; t.Cells[1, 2].Value = "V"; t.Cells[1, 3].Value = "D";
            for (int r = 2; r <= 21; r++) { t.Cells[r, 1].Value = "k" + r; t.Cells[r, 2].Value = r * 1.5; t.Cells[r, 3].Value = new DateTime(2026, 1, 1).AddDays(r); }
            t.Cells[2, 3, 21, 3].Style.Numberformat.Format = "yyyy-mm-dd hh:mm:ss";
            var tbl = t.Tables.Add(t.Cells[1, 1, 21, 3], "T1_2");
            tbl.TableStyle = TableStyles.Medium2;
            t.Cells[1, 1, 1, 3].Style.Font.Bold = true;
            t.Cells[1, 1, 21, 3].AutoFitColumns(4, 80);
            var ch = t.Drawings.AddChart("Chart1_1", eChartType.ColumnClustered);
            ch.Series.Add(ExcelCellBase.GetFullAddress("Table", "B2:B21"), ExcelCellBase.GetFullAddress("Table", "A2:A21"));
            ch.SetPosition(1, 0, 5, 0); ch.SetSize(400, 300);
            p.SaveAs(new FileInfo(a_path));
        }
    }

}
