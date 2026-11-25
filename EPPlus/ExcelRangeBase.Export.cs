/*******************************************************************************
 * orabooster-epplus: this file is new in the fork (issue #32, item B6 of issue #8). It is not part of EPPlus 4.5 or EPPlusFree
 * and has no code from EPPlus 5 or later. Written by OraScope Co.,Ltd. 2026, under the license of the library (LGPL-3.0).
 *
 * The export direction of a range: LoadFromDataTable and LoadFromText put data into a range, the methods here take it out.
 *   ToDataTable          a DataTable, typed columns
 *   ToText, SaveToText   delimited text (RFC 4180), the reverse of LoadFromText with the same ExcelTextFormat
 *   ToJson               a JSON array (RFC 8259)
 *   ToHtml               an HTML table with the text of the cells as Excel shows it
 * Nothing here changes the workbook.
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml
{
    public partial class ExcelRangeBase
    {
        #region The cells of the export

        /// <summary>
        /// The rows and columns that are exported: the first area of the range, cut off at the last row and column of the sheet that
        /// has a cell (so that "A:C" is not a million rows). False when there is nothing.
        /// </summary>
        private bool GetExportArea(out int fromRow, out int fromCol, out int toRow, out int toCol)
        {
            fromRow = _fromRow; fromCol = _fromCol; toRow = _toRow; toCol = _toCol;
            ExcelAddressBase dimension = _worksheet.Dimension;
            if (dimension == null) return false;
            if (toRow > dimension.End.Row) toRow = dimension.End.Row;
            if (toCol > dimension.End.Column) toCol = dimension.End.Column;
            return fromRow <= toRow && fromCol <= toCol;
        }

        /// <summary>Reads the values of the export: the value of a cell, with a number in a date format as a date.</summary>
        private sealed class ExportReader
        {
            readonly ExcelWorksheet _ws;
            readonly Dictionary<int, bool> _dateStyles = new Dictionary<int, bool>();

            internal ExportReader(ExcelWorksheet ws) { _ws = ws; }

            internal object Value(int row, int col)
            {
                object v = _ws.GetValue(row, col); // the text of a rich text cell, not its XML
                if (v == null) return null;
                if (v is string) return v;
                if (v is DateTime || v is bool || v is TimeSpan) return v;
                if (v is ExcelErrorValue) return v.ToString();
                if (v is double || v is float || v is decimal || v is int || v is long || v is short || v is byte
                    || v is uint || v is ulong || v is ushort || v is sbyte)
                {
                    double d = Convert.ToDouble(v, CultureInfo.InvariantCulture);
                    // a number in a cell with a date format is a date (cells of a loaded workbook already hold a DateTime)
                    if (d > -657435.0 && d < 2958466.0 && IsDateStyle(row, col)) return DateTime.FromOADate(d);
                    return d;
                }
                return v;
            }

            bool IsDateStyle(int row, int col)
            {
                // the style of the cell, else of its row, else of its column (as ExcelRangeBase.Style finds it)
                int styleId = 0;
                if (!_ws.ExistsStyleInner(row, col, ref styleId) && !_ws.ExistsStyleInner(row, 0, ref styleId))
                {
                    ExcelColumn column = _ws.GetColumn(col);
                    styleId = column == null ? 0 : column.StyleID;
                }
                bool isDate;
                if (_dateStyles.TryGetValue(styleId, out isDate)) return isDate;
                ExcelStyles styles = _ws.Workbook.Styles;
                int formatId = styles.CellXfs[styleId].NumberFormatId;
                isDate = false;
                for (int i = 0; i < styles.NumberFormats.Count; i++)
                {
                    if (styles.NumberFormats[i].NumFmtId == formatId)
                    {
                        isDate = styles.NumberFormats[i].FormatTranslator.DataType == ExcelNumberFormatXml.eFormatType.DateTime;
                        break;
                    }
                }
                _dateStyles.Add(styleId, isDate);
                return isDate;
            }
        }

        /// <summary>
        /// The names of the columns of the export: the text of the header row when there is one, else Column1, Column2 and so on.
        /// An empty header gets the name Column and its position. A name that was used before gets _2, _3 and so on.
        /// </summary>
        private static string[] GetExportNames(ExportReader reader, bool firstRowIsHeader, int headerRow, int fromCol, int toCol)
        {
            string[] names = new string[toCol - fromCol + 1];
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase); // DataTable column names do not tell case apart
            for (int c = fromCol; c <= toCol; c++)
            {
                string name = null;
                if (firstRowIsHeader)
                {
                    object v = reader.Value(headerRow, c);
                    if (v != null) name = Convert.ToString(v, CultureInfo.InvariantCulture).Trim();
                }
                if (string.IsNullOrEmpty(name)) name = "Column" + (c - fromCol + 1).ToString(CultureInfo.InvariantCulture);
                string unique = name;
                for (int n = 2; !used.Add(unique); n++) unique = name + "_" + n.ToString(CultureInfo.InvariantCulture);
                names[c - fromCol] = unique;
            }
            return names;
        }

        private static string IsoDate(DateTime d)
        {
            if (d.TimeOfDay.Ticks == 0) return d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            if (d.Millisecond == 0) return d.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            return d.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
        }

        #endregion

        #region ToDataTable

        /// <summary>
        /// The cells of the range as a <see cref="DataTable"/>: one column for every column of the range and one row for every row.
        /// A range of whole rows or columns ends at the last row and column of the sheet that has a cell. If the range has
        /// several areas, the first one is used.
        /// </summary>
        /// <param name="firstRowIsHeader">
        /// True: the first row of the range names the columns and is not a row of the table. An empty header is named Column and its
        /// position; a name that comes again gets _2, _3 and so on. False: the columns are named Column1, Column2 and so on.
        /// </param>
        /// <returns>
        /// The table, named after the worksheet. A column is typed when all its values that are not empty have one type:
        /// <see cref="double"/> for numbers, <see cref="DateTime"/> (also for numbers in a date format), <see cref="bool"/>,
        /// <see cref="string"/> or <see cref="TimeSpan"/>. A column with values of several types is <see cref="object"/>, a column
        /// with no values is <see cref="string"/>. An empty cell is <see cref="DBNull"/>, an error value is its text (#DIV/0!).
        /// </returns>
        public DataTable ToDataTable(bool firstRowIsHeader = true)
        {
            var table = new DataTable(_worksheet.Name) { Locale = CultureInfo.InvariantCulture };
            int fromRow, fromCol, toRow, toCol;
            if (!GetExportArea(out fromRow, out fromCol, out toRow, out toCol)) return table;
            var reader = new ExportReader(_worksheet);
            string[] names = GetExportNames(reader, firstRowIsHeader, fromRow, fromCol, toCol);
            int firstData = firstRowIsHeader ? fromRow + 1 : fromRow;
            int columns = toCol - fromCol + 1, rows = Math.Max(0, toRow - firstData + 1);

            object[][] values = new object[rows][];
            Type[] types = new Type[columns];
            bool[] mixed = new bool[columns];
            for (int r = 0; r < rows; r++)
            {
                object[] line = values[r] = new object[columns];
                for (int c = 0; c < columns; c++)
                {
                    object v = reader.Value(firstData + r, fromCol + c);
                    if (v == null) { line[c] = DBNull.Value; continue; }
                    line[c] = v;
                    Type t = v.GetType();
                    if (types[c] == null) types[c] = t;
                    else if (types[c] != t) mixed[c] = true;
                }
            }
            for (int c = 0; c < columns; c++)
            {
                table.Columns.Add(names[c], mixed[c] ? typeof(object) : (types[c] ?? typeof(string)));
            }
            table.BeginLoadData();
            for (int r = 0; r < rows; r++) table.Rows.Add(values[r]);
            table.EndLoadData();
            table.AcceptChanges();
            return table;
        }

        #endregion

        #region ToText, SaveToText

        /// <summary>
        /// The cells of the range as delimited text (CSV), one line for every row: the reverse of
        /// <see cref="LoadFromText(string, ExcelTextFormat)"/>.
        /// </summary>
        /// <param name="format">
        /// The delimiter, the text qualifier, the end of a line and the culture of the numbers. Null is the default
        /// <see cref="ExcelTextFormat"/>: a comma, lines that end with CR LF, the invariant culture.
        /// <para>The data types and the lines to skip are for loading and are not used here.</para>
        /// </param>
        /// <returns>
        /// The text. The last line also ends with the end of a line.
        /// <para>Numbers are written with the culture of the format and with all their digits, dates as yyyy-MM-dd or
        /// yyyy-MM-dd HH:mm:ss (which every culture reads), booleans as TRUE and FALSE, empty cells as nothing.</para>
        /// <para>A value is put between text qualifiers when it has the delimiter, the qualifier (which is then doubled), a line
        /// break or a space at an end in it, or when it is a text that would be read as a number or a date (00123, 1e5, 2026-01-02).
        /// If the format has no text qualifier, the double quote is used. To load such text again, set
        /// <see cref="ExcelTextFormat.TextQualifier"/> to that character.</para>
        /// </returns>
        public string ToText(ExcelTextFormat format = null)
        {
            if (format == null) format = new ExcelTextFormat();
            CultureInfo culture = format.Culture ?? CultureInfo.InvariantCulture;
            char delimiter = format.Delimiter;
            char qualifier = format.TextQualifier == '\0' ? '"' : format.TextQualifier;
            string eol = string.IsNullOrEmpty(format.EOL) ? "\r\n" : format.EOL;

            var sb = new StringBuilder();
            int fromRow, fromCol, toRow, toCol;
            if (!GetExportArea(out fromRow, out fromCol, out toRow, out toCol)) return "";
            var reader = new ExportReader(_worksheet);
            for (int r = fromRow; r <= toRow; r++)
            {
                for (int c = fromCol; c <= toCol; c++)
                {
                    if (c > fromCol) sb.Append(delimiter);
                    object v = reader.Value(r, c);
                    if (v == null) continue;
                    string text;
                    bool isText = false;
                    if (v is string) { text = (string)v; isText = true; }
                    else if (v is double) text = ((double)v).ToString("R", culture);
                    else if (v is DateTime) text = IsoDate((DateTime)v);
                    else if (v is bool) text = (bool)v ? "TRUE" : "FALSE";
                    else if (v is IFormattable) text = ((IFormattable)v).ToString(null, culture);
                    else { text = v.ToString(); isText = true; }
                    if (NeedsQualifier(text, isText, delimiter, qualifier, culture))
                    {
                        sb.Append(qualifier).Append(text.Replace(qualifier.ToString(), new string(qualifier, 2))).Append(qualifier);
                    }
                    else
                    {
                        sb.Append(text);
                    }
                }
                sb.Append(eol);
            }
            return sb.ToString();
        }

        private static bool NeedsQualifier(string text, bool isText, char delimiter, char qualifier, CultureInfo culture)
        {
            if (text.Length == 0) return false;
            if (text.IndexOf(delimiter) >= 0 || text.IndexOf(qualifier) >= 0 || text.IndexOf('\r') >= 0 || text.IndexOf('\n') >= 0) return true;
            if (!isText) return false;
            if (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1])) return true;
            // a text that LoadFromText would turn into a number or a date (it tries both for a value that is not qualified)
            string number = text.EndsWith("%", StringComparison.Ordinal) ? text.Substring(0, text.Length - 1) : text;
            double d;
            if (double.TryParse(number, NumberStyles.Any, culture, out d)) return true;
            DateTime dt;
            return DateTime.TryParse(text, culture, DateTimeStyles.None, out dt);
        }

        /// <summary>
        /// Writes the cells of the range to a file of delimited text (CSV). See <see cref="ToText(ExcelTextFormat)"/>.
        /// </summary>
        /// <param name="file">The file. It is made new.</param>
        /// <param name="format">
        /// Null: the default format, and the file is UTF-8 with a byte order mark, which Excel reads right.
        /// Else the file is written with <see cref="ExcelTextFormat.Encoding"/> of the format. Mind that the encoding of a new
        /// <see cref="ExcelTextFormat"/> is ASCII, which has no letters outside English.
        /// </param>
        public void SaveToText(FileInfo file, ExcelTextFormat format = null)
        {
            if (file == null) throw new ArgumentNullException("file");
            Encoding encoding = format == null ? new UTF8Encoding(true) : (format.Encoding ?? new UTF8Encoding(true));
            File.WriteAllText(file.FullName, ToText(format), encoding);
        }

        #endregion

        #region ToJson

        /// <summary>
        /// The cells of the range as JSON (RFC 8259), an array with one element for every row.
        /// </summary>
        /// <param name="firstRowIsHeader">
        /// True: every row is an object whose members are named by the first row of the range (named as in
        /// <see cref="ToDataTable(bool)"/>). False: every row is an array of its values.
        /// </param>
        /// <returns>
        /// The JSON text, with no line breaks. Numbers are JSON numbers with all their digits, booleans true and false, empty cells
        /// null, dates text as yyyy-MM-ddTHH:mm:ss (with .fff when there are milliseconds), anything else text.
        /// </returns>
        public string ToJson(bool firstRowIsHeader = true)
        {
            int fromRow, fromCol, toRow, toCol;
            if (!GetExportArea(out fromRow, out fromCol, out toRow, out toCol)) return "[]";
            var reader = new ExportReader(_worksheet);
            string[] names = firstRowIsHeader ? GetExportNames(reader, true, fromRow, fromCol, toCol) : null;
            string[] quotedNames = null;
            if (names != null)
            {
                quotedNames = new string[names.Length];
                var one = new StringBuilder();
                for (int i = 0; i < names.Length; i++) { one.Length = 0; AppendJsonString(one, names[i]); quotedNames[i] = one.ToString(); }
            }
            var sb = new StringBuilder();
            sb.Append('[');
            int firstData = firstRowIsHeader ? fromRow + 1 : fromRow;
            for (int r = firstData; r <= toRow; r++)
            {
                if (r > firstData) sb.Append(',');
                sb.Append(names != null ? '{' : '[');
                for (int c = fromCol; c <= toCol; c++)
                {
                    if (c > fromCol) sb.Append(',');
                    if (quotedNames != null) sb.Append(quotedNames[c - fromCol]).Append(':');
                    AppendJsonValue(sb, reader.Value(r, c));
                }
                sb.Append(names != null ? '}' : ']');
            }
            sb.Append(']');
            return sb.ToString();
        }

        private static void AppendJsonValue(StringBuilder sb, object v)
        {
            if (v == null) { sb.Append("null"); return; }
            if (v is string) { AppendJsonString(sb, (string)v); return; }
            if (v is double)
            {
                double d = (double)v;
                if (double.IsNaN(d) || double.IsInfinity(d)) sb.Append("null"); // JSON has no such numbers
                else sb.Append(d.ToString("R", CultureInfo.InvariantCulture));
                return;
            }
            if (v is bool) { sb.Append((bool)v ? "true" : "false"); return; }
            if (v is DateTime)
            {
                DateTime d = (DateTime)v;
                sb.Append('"').Append(d.ToString(d.Millisecond == 0 ? "yyyy-MM-ddTHH:mm:ss" : "yyyy-MM-ddTHH:mm:ss.fff", CultureInfo.InvariantCulture)).Append('"');
                return;
            }
            AppendJsonString(sb, Convert.ToString(v, CultureInfo.InvariantCulture));
        }

        private static void AppendJsonString(StringBuilder sb, string s)
        {
            sb.Append('"');
            foreach (char ch in s)
            {
                switch (ch)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    case '\b': sb.Append("\\b"); break;
                    case '\f': sb.Append("\\f"); break;
                    default:
                        if (ch < ' ') sb.Append("\\u").Append(((int)ch).ToString("x4", CultureInfo.InvariantCulture));
                        else sb.Append(ch);
                        break;
                }
            }
            sb.Append('"');
        }

        #endregion

        #region ToHtml

        /// <summary>
        /// The cells of the range as an HTML table, with the text of every cell as Excel shows it (its number format applied).
        /// </summary>
        /// <param name="firstRowIsHeader">True: the first row of the range is the head of the table (thead, th).</param>
        /// <returns>
        /// A table element, one line for every row. Text is escaped and a line break in a cell is a br element.
        /// A merged cell is one cell with colspan and rowspan (inside the range). A cell with a number or a date is aligned to the
        /// right, unless the cell has its own horizontal alignment; bold and italic cells are bold and italic. No other format
        /// (colours, borders, widths) is written: a style sheet can be put on the table.
        /// </returns>
        public string ToHtml(bool firstRowIsHeader = true)
        {
            var sb = new StringBuilder();
            int fromRow, fromCol, toRow, toCol;
            if (!GetExportArea(out fromRow, out fromCol, out toRow, out toCol)) return "<table>\r\n</table>";
            sb.Append("<table>\r\n");
            if (firstRowIsHeader)
            {
                sb.Append("<thead>\r\n");
                AppendHtmlRow(sb, fromRow, fromRow, fromCol, fromRow, toCol, "th"); // the head is one row
                sb.Append("</thead>\r\n");
            }
            sb.Append("<tbody>\r\n");
            int firstBody = firstRowIsHeader ? fromRow + 1 : fromRow;
            for (int r = firstBody; r <= toRow; r++)
            {
                AppendHtmlRow(sb, r, firstBody, fromCol, toRow, toCol, "td");
            }
            sb.Append("</tbody>\r\n</table>");
            return sb.ToString();
        }

        /// <summary>One row. fromRow and toRow are the rows of its part of the table (the head or the body): a merged area spans inside it.</summary>
        private void AppendHtmlRow(StringBuilder sb, int row, int fromRow, int fromCol, int toRow, int toCol, string tag)
        {
            sb.Append("<tr>");
            for (int c = fromCol; c <= toCol; c++)
            {
                int colSpan = 1, rowSpan = 1;
                string merged = _worksheet.MergedCells[row, c];
                if (merged != null)
                {
                    var area = new ExcelAddressBase(merged);
                    // only the first cell of a merged area that is inside this part of the table is written; it spans what is inside.
                    // An area that goes from the head into the body is a cell of the head and a cell of the body.
                    int firstRow = Math.Max(area.Start.Row, fromRow), firstCol = Math.Max(area.Start.Column, fromCol);
                    if (row != firstRow || c != firstCol) continue;
                    colSpan = Math.Min(area.End.Column, toCol) - firstCol + 1;
                    rowSpan = Math.Min(area.End.Row, toRow) - firstRow + 1;
                    // the value and the style of a merged area are those of its first cell
                    AppendHtmlCell(sb, area.Start.Row, area.Start.Column, tag, colSpan, rowSpan);
                    continue;
                }
                AppendHtmlCell(sb, row, c, tag, colSpan, rowSpan);
            }
            sb.Append("</tr>\r\n");
        }

        private void AppendHtmlCell(StringBuilder sb, int row, int col, string tag, int colSpan, int rowSpan)
        {
            ExcelRange cell = _worksheet.Cells[row, col];
            object v = cell.Value;
            sb.Append('<').Append(tag);
            if (colSpan > 1) sb.Append(" colspan=\"").Append(colSpan.ToString(CultureInfo.InvariantCulture)).Append('"');
            if (rowSpan > 1) sb.Append(" rowspan=\"").Append(rowSpan.ToString(CultureInfo.InvariantCulture)).Append('"');
            string style = "";
            if (v != null)
            {
                ExcelStyle cellStyle = cell.Style;
                switch (cellStyle.HorizontalAlignment)
                {
                    case ExcelHorizontalAlignment.Left: style += "text-align:left;"; break;
                    case ExcelHorizontalAlignment.Center:
                    case ExcelHorizontalAlignment.CenterContinuous: style += "text-align:center;"; break;
                    case ExcelHorizontalAlignment.Right: style += "text-align:right;"; break;
                    default:
                        // General: Excel puts numbers and dates to the right
                        if (!(v is string) && !(v is bool) && !(v is ExcelErrorValue)) style += "text-align:right;";
                        break;
                }
                if (cellStyle.Font.Bold && tag != "th") style += "font-weight:bold;";
                if (cellStyle.Font.Italic) style += "font-style:italic;";
            }
            if (style.Length > 0) sb.Append(" style=\"").Append(style).Append('"');
            sb.Append('>');
            if (v != null) AppendHtmlText(sb, cell.Text);
            sb.Append("</").Append(tag).Append('>');
        }

        private static void AppendHtmlText(StringBuilder sb, string text)
        {
            foreach (char ch in text)
            {
                switch (ch)
                {
                    case '&': sb.Append("&amp;"); break;
                    case '<': sb.Append("&lt;"); break;
                    case '>': sb.Append("&gt;"); break;
                    case '"': sb.Append("&quot;"); break;
                    case '\r': break;
                    case '\n': sb.Append("<br>"); break;
                    default: sb.Append(ch); break;
                }
            }
        }

        #endregion
    }
}
