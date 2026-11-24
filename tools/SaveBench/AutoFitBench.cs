// AutoFitColumns: how long it takes, and how the widths compare with Excel's own autofit (issue #14).
//   SaveBench autofit [rows] [cols] [runs]    time AutoFitColumns over every row
//   SaveBench autofitsample <path>            write a workbook with many kinds of text, widths set by AutoFitColumns;
//                                             autofit-in-excel.ps1 then compares each width with Excel's autofit
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using OfficeOpenXml;

static class AutoFitBench
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

    internal static void Run(int rows, int cols, int runs)
    {
        var data = new List<object[]>(rows + 1);
        var head = new object[cols];
        for (int c = 0; c < cols; c++) head[c] = "COLUMN_" + (c + 1);
        data.Add(head);
        for (int r = 0; r < rows; r++) { var a = new object[cols]; for (int c = 0; c < cols; c++) a[c] = Value(r + 2, c + 1); data.Add(a); }

        for (int run = 0; run < runs; run++)
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("S");
                ws.Cells[1, 1].LoadFromArrays(data);
                for (int c = 1; c <= cols; c++) if (c % 5 == 2) ws.Cells[2, c, rows + 1, c].Style.Numberformat.Format = "yyyy-mm-dd hh:mm:ss";
                ws.Cells[1, 1, 1, cols].Style.Font.Bold = true;
                var sw = Stopwatch.StartNew();
                ws.Cells[1, 1, rows + 1, cols].AutoFitColumns();
                long ms = sw.ElapsedMilliseconds;
                var widths = new List<string>();
                for (int c = 1; c <= Math.Min(cols, 5); c++) widths.Add(ws.Column(c).Width.ToString("0.###"));
                Console.WriteLine("autofit {0} rows x {1} cols: {2} ms   widths {3}", rows, cols, ms, string.Join(" ", widths));
            }
        }
    }

    // One column per kind of text. Row 1 is a bold header, the rows below hold the values.
    internal static void WriteSample(string path)
    {
        var columns = new List<KeyValuePair<string, object[]>>
        {
            Col("short", "abc", "de"),
            Col("x40", new string('x', 40)),
            Col("sentence", "The quick brown fox jumps over the lazy dog", "select * from dual"),
            Col("int", 1, 123456789, 42),
            Col("decimal", 1234567.891, 0.5),
            Col("date", new DateTime(2026, 11, 24, 13, 45, 59), new DateTime(2026, 1, 2)),
            Col("korean", "한글 열 이름과 값", "세션 모니터링"),
            Col("wide", "WWWWWWWWWWWWWWWWWWWW", "MMMMMMMMMM"),
            Col("narrow", "iiiiiiiiiiiiiiiiiiiiiiiiiiiiii", "llllllllll"),
            Col("A_LONG_BOLD_HEADER_NAME_OF_A_COLUMN", 1),
            Col("big font", "Text in 16 points", "more"),
            Col("sql", "SELECT /*+ INDEX(t ix_t_01) */ owner, object_name, object_type, created FROM dba_objects t WHERE owner = :b1 ORDER BY created DESC"),
            Col("digits as text", "0000012345", "00042"),
            Col("upper", "TABLESPACE_NAME", "SYSAUX", "UNDOTBS1"),
            Col("mixed", "SQL_ID 7x9k2m4p1q8zr / plan 3128390112", "PX COORDINATOR"),
            Col("spaces", "a  b   c    d     e", " leading"),
            Col("symbols", "%_#@!$&*()[]{}<>+=-/\\|", "50.25%"),
            Col("negative", -1234.5, -0.001),
        };
        using (var p = new ExcelPackage())
        {
            var ws = p.Workbook.Worksheets.Add("S");
            int rows = 1;
            for (int c = 0; c < columns.Count; c++)
            {
                ws.Cells[1, c + 1].Value = columns[c].Key;
                object[] values = columns[c].Value;
                for (int r = 0; r < values.Length; r++) ws.Cells[r + 2, c + 1].Value = values[r];
                rows = Math.Max(rows, values.Length + 1);
            }
            ws.Cells[1, 1, 1, columns.Count].Style.Font.Bold = true;
            ws.Cells[2, 5, rows, 5].Style.Numberformat.Format = "#,##0.00";
            ws.Cells[2, 6, rows, 6].Style.Numberformat.Format = "yyyy-mm-dd hh:mm:ss";
            ws.Cells[2, 11, rows, 11].Style.Font.Size = 16;
            ws.Cells[2, 18, rows, 18].Style.Numberformat.Format = "#,##0.000";
            ws.Cells[1, 1, rows, columns.Count].AutoFitColumns();
            p.SaveAs(new FileInfo(path));
        }
    }

    // A workbook of random cells with the things AutoFitColumns looks at: texts of many shapes, numbers and dates with
    // formats, fonts, indents, rotated, wrapped and merged cells, a hidden column, a table with filter buttons, a width limit
    // and a range of several addresses. Written with two builds of the library, the two files must be the same.
    internal static void WriteRandom(string path, int seed)
    {
        var rnd = new Random(seed);
        string ascii = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 _-.,:;/%#()[]+*'\"!?@&=<>|\\{}~^$`";
        string korean = "가나다라마바사아자차카타파하세션모니터링실행계획테이블공간대기이벤트";
        // fonts with digits of one width and of several widths (Georgia, Gabriola), one that is not installed
        string[] fonts = { "Arial", "Consolas", "Georgia", "Times New Roman", "Segoe UI", "Malgun Gothic", "Courier New", "Verdana", "Gabriola", "No Such Font" };
        string[] formats = { null, "#,##0.00", "0", "0.000", "#,##0", "0.0%", "yyyy-mm-dd", "yyyy-mm-dd hh:mm:ss", "hh:mm:ss", "@", "0.00E+00", "#,##0.00;[Red]-#,##0.00" };
        using (var p = new ExcelPackage())
        {
            for (int s = 0; s < 3; s++)
            {
                var ws = p.Workbook.Worksheets.Add("S" + s);
                // enough rows for a font to get past the point where its digits are checked (see AutoFitFont in ExcelRangeBase)
                int rows = 1500 + rnd.Next(1000), cols = 24;
                for (int c = 1; c <= cols; c++)
                {
                    ws.Cells[1, c].Value = "HEAD_" + c + new string('X', rnd.Next(12));
                    int kind = rnd.Next(8);
                    string format = formats[rnd.Next(formats.Length)];
                    // a date format on a number that is not a date makes AutoFitColumns throw (issue #15), so dates get the date formats and nothing else does
                    bool dateFormat = format != null && (format.Contains("yy") || format.Contains("hh"));
                    if (dateFormat != (kind == 2)) format = kind == 2 ? "yyyy-mm-dd hh:mm:ss" : null;
                    for (int r = 2; r <= rows; r++)
                    {
                        if (rnd.Next(12) == 0) continue;
                        object v;
                        switch (kind)
                        {
                            case 0: v = rnd.Next(1000000); break;
                            case 1: v = Math.Round((rnd.NextDouble() - 0.3) * Math.Pow(10, rnd.Next(9)), rnd.Next(5)); break;
                            case 2: v = new DateTime(2020, 1, 1).AddSeconds(rnd.Next(200000000)); break;
                            case 3: { var sb = new System.Text.StringBuilder(); int n = rnd.Next(40); for (int i = 0; i < n; i++) sb.Append(ascii[rnd.Next(ascii.Length)]); v = sb.ToString(); break; }
                            case 4: { var sb = new System.Text.StringBuilder(); int n = 1 + rnd.Next(14); for (int i = 0; i < n; i++) sb.Append(rnd.Next(4) == 0 ? ascii[rnd.Next(ascii.Length)] : korean[rnd.Next(korean.Length)]); v = sb.ToString(); break; }
                            case 5: v = "SQL_" + rnd.Next().ToString("X8"); break;
                            case 6: v = rnd.Next(3) == 0 ? (object)(rnd.Next(100000).ToString("D7")) : (object)("id " + rnd.Next(50) + "." + rnd.Next(10) + ":" + rnd.Next(60) + "-" + rnd.Next(9) + "," + rnd.Next(9)); break;
                            default: v = rnd.Next(2) == 0 ? (object)((decimal)rnd.Next(100000) / 100m) : (object)(rnd.Next(2) == 0); break;
                        }
                        ws.Cells[r, c].Value = v;
                    }
                    if (format != null) ws.Cells[2, c, rows, c].Style.Numberformat.Format = format;
                    switch (rnd.Next(9))
                    {
                        case 0: ws.Cells[2, c, rows, c].Style.Font.Bold = true; break;
                        case 1: ws.Cells[2, c, rows, c].Style.Font.Italic = true; break;
                        case 2: ws.Cells[2, c, rows, c].Style.Font.Size = 8 + rnd.Next(10); break;
                        case 3: ws.Cells[2, c, rows, c].Style.Font.Name = fonts[rnd.Next(fonts.Length)]; break;
                        case 4: ws.Cells[2, c, rows, c].Style.Indent = 1 + rnd.Next(4); break;
                        case 5: ws.Cells[2 + rnd.Next(20), c].Style.TextRotation = rnd.Next(181); break;
                        case 6: ws.Cells[2 + rnd.Next(20), c].Style.WrapText = true; break;
                    }
                }
                ws.Cells[1, 1, 1, cols].Style.Font.Bold = true;
                // a style on a whole row and on a whole column, a hidden column, merged cells
                ws.Row(5).Style.Font.Size = 14;
                ws.Column(7).Style.Font.Italic = true;
                ws.Column(9).Hidden = true;
                ws.Cells[10, 3, 10, 5].Merge = true;
                ws.Cells[30, 12, 32, 12].Merge = true;
                ws.Cells[40, 2].Value = new string('W', 300);     // longer than a text that is kept
                ws.Cells[41, 2].Value = "line one\nline two is longer";
                if (s != 1) ws.Tables.Add(ws.Cells[1, 1, rows, 10], "Table_" + s);
                else ws.Cells[1, 12, rows, 20].AutoFilter = true;

                switch (s)
                {
                    case 0: ws.Cells[1, 1, rows, cols].AutoFitColumns(); break;
                    case 1: ws.Cells[1, 1, rows, cols].AutoFitColumns(4, 30); break;
                    default: ws.Cells["A1:D200,F5:M300,O1:X" + rows].AutoFitColumns(2); break;
                }
            }
            p.SaveAs(new FileInfo(path));
        }
    }

    static KeyValuePair<string, object[]> Col(string head, params object[] values)
    {
        return new KeyValuePair<string, object[]>(head, values);
    }
}
