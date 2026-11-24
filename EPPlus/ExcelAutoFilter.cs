// orabooster-epplus: filter criteria and sort state of an autofilter (https://github.com/orascope/orabooster-epplus/issues/17).
// Written from ECMA-376 Part 1, 18.3.2 (autoFilter, filterColumn, filters, customFilters) and 18.3.1 (sortState),
// and compared with files in which Excel itself set the same filters.
// (c) 2026 OraScope Co.,Ltd. GNU Lesser General Public License v3.0, as the rest of the library.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using OfficeOpenXml.Table;

namespace OfficeOpenXml
{
    /// <summary>
    /// How a cell is compared with the value of a filter condition.
    /// </summary>
    public enum eFilterOperator
    {
        /// <summary>Equal. For a text the value can hold * (any characters) and ? (one character); ~ before one of them means the character itself.</summary>
        Equal,
        /// <summary>Not equal. Wildcards as for Equal.</summary>
        NotEqual,
        /// <summary>Greater than</summary>
        GreaterThan,
        /// <summary>Greater than or equal</summary>
        GreaterThanOrEqual,
        /// <summary>Less than</summary>
        LessThan,
        /// <summary>Less than or equal</summary>
        LessThanOrEqual
    }

    /// <summary>
    /// The filter buttons of a worksheet range (ExcelWorksheet.AutoFilter) or of a table (ExcelTable.AutoFilter):
    /// what each column shows, and how the rows are sorted.
    /// Excel does not apply a filter when it opens a file. The rows that do not pass are hidden here, when the package is
    /// saved or when Apply is called.
    /// </summary>
    public sealed class ExcelAutoFilter
    {
        static readonly string[] s_operators = { "equal", "notEqual", "greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual" };

        readonly ExcelWorksheet _ws;
        readonly ExcelTable _table;   // null for the filter of the worksheet
        bool _changed;

        internal ExcelAutoFilter(ExcelWorksheet ws, ExcelTable table)
        {
            _ws = ws;
            _table = table;
        }

        XmlNamespaceManager Ns { get { return _ws.NameSpaceManager; } }
        XmlNode Owner { get { return _table == null ? _ws.TopNode : _table.TopNode; } }

        XmlElement Node
        {
            get
            {
                var node = Owner.SelectSingleNode("d:autoFilter", Ns) as XmlElement;
                if (node == null) throw new InvalidOperationException("The autofilter was removed");
                return node;
            }
        }

        /// <summary>
        /// The cells of the autofilter: the row with the buttons and the rows below it.
        /// </summary>
        public ExcelAddressBase Address
        {
            get { return new ExcelAddressBase(Node.GetAttribute("ref")); }
        }

        /// <summary>
        /// Shows the rows whose cell in the column has one of the texts, as the cell shows it: with its number format, in the
        /// culture of the thread (a 12 with the format 0.00 is "12.00", or "12,00" where the decimal sign is a comma).
        /// Upper and lower case are the same. A null or empty text stands for the empty cells.
        /// </summary>
        /// <param name="column">The column within the autofilter, zero based</param>
        /// <param name="values">The texts to show</param>
        public void ShowValues(int column, params string[] values)
        {
            CheckColumn(column);
            if (values == null || values.Length == 0) throw new ArgumentException("No value to show", "values");

            XmlElement filters = NewFilterColumn(column, "filters");
            foreach (string v in values)
            {
                if (string.IsNullOrEmpty(v))
                {
                    filters.SetAttribute("blank", "1");
                }
                else
                {
                    XmlElement f = filters.OwnerDocument.CreateElement("filter", ExcelPackage.schemaMain);
                    f.SetAttribute("val", v);
                    filters.AppendChild(f);
                }
            }
            _changed = true;
        }

        /// <summary>
        /// Shows the rows whose cell in the column meets the condition.
        /// A number or a date is compared with the number of the cell; a text is compared with the text the cell shows, without case.
        /// A null or empty value stands for the empty cells: Equal shows them, NotEqual shows the cells that are not empty.
        /// A number cell does not match a text with wildcards and is not greater or less than a text.
        /// </summary>
        /// <remarks>
        /// Excel compares "equal" with the text the cell shows, so for Equal with a number or a date this writes
        /// "not less and not greater", which compares the number whatever the number format is.
        /// </remarks>
        /// <param name="column">The column within the autofilter, zero based</param>
        /// <param name="op">How the cell is compared with the value</param>
        /// <param name="value">A number, a date or a text</param>
        public void ShowWhere(int column, eFilterOperator op, object value)
        {
            CheckColumn(column);
            XmlElement custom = NewFilterColumn(column, "customFilters");
            if (op == eFilterOperator.Equal && IsNumber(value))
            {
                // Excel compares "equal" with the text the cell shows, so "equal to 12" would miss a 12 that is shown as 12.00.
                // Not less and not greater compares the number.
                custom.SetAttribute("and", "1");
                AddCustomFilter(custom, eFilterOperator.GreaterThanOrEqual, value);
                AddCustomFilter(custom, eFilterOperator.LessThanOrEqual, value);
            }
            else
            {
                AddCustomFilter(custom, op, value);
            }
            _changed = true;
        }

        static bool IsNumber(object value)
        {
            return value != null && !(value is string) && !(value is bool) && !(value is char) && (value is DateTime || value is TimeSpan || value is decimal || value.GetType().IsPrimitive);
        }

        /// <summary>
        /// Shows the rows whose cell in the column meets two conditions: both (and) or at least one (or).
        /// </summary>
        /// <remarks>
        /// Here Equal is written as it is, and Excel compares it with the text the cell shows: Equal to 12 does not show a 12
        /// that is shown as 12.00. NotEqual and the greater and less operators compare the number.
        /// </remarks>
        /// <param name="column">The column within the autofilter, zero based</param>
        /// <param name="op1">How the cell is compared with the first value</param>
        /// <param name="value1">A number, a date or a text</param>
        /// <param name="and">True: both conditions. False: one of them is enough.</param>
        /// <param name="op2">How the cell is compared with the second value</param>
        /// <param name="value2">A number, a date or a text</param>
        public void ShowWhere(int column, eFilterOperator op1, object value1, bool and, eFilterOperator op2, object value2)
        {
            CheckColumn(column);
            XmlElement custom = NewFilterColumn(column, "customFilters");
            if (and) custom.SetAttribute("and", "1");
            AddCustomFilter(custom, op1, value1);
            AddCustomFilter(custom, op2, value2);
            _changed = true;
        }

        /// <summary>
        /// Takes the filter off a column. Its rows show again when the package is saved or Apply is called.
        /// </summary>
        /// <param name="column">The column within the autofilter, zero based</param>
        public void Clear(int column)
        {
            CheckColumn(column);
            XmlNode old = Node.SelectSingleNode("d:filterColumn[@colId='" + column.ToString(CultureInfo.InvariantCulture) + "']", Ns);
            if (old != null)
            {
                Node.RemoveChild(old);
                _changed = true;
            }
        }

        /// <summary>
        /// Takes the filter off every column.
        /// </summary>
        public void Clear()
        {
            foreach (XmlNode old in Node.SelectNodes("d:filterColumn", Ns))
            {
                Node.RemoveChild(old);
                _changed = true;
            }
        }

        /// <summary>
        /// True when the column has a filter.
        /// </summary>
        /// <param name="column">The column within the autofilter, zero based</param>
        public bool IsFiltered(int column)
        {
            CheckColumn(column);
            return Node.SelectSingleNode("d:filterColumn[@colId='" + column.ToString(CultureInfo.InvariantCulture) + "']", Ns) != null;
        }

        /// <summary>
        /// Sorts the rows below the buttons by a column and records the sort, so that the button of the column shows the arrow.
        /// </summary>
        /// <param name="column">The column within the autofilter, zero based</param>
        /// <param name="descending">Largest first</param>
        public void Sort(int column, bool descending = false)
        {
            CheckColumn(column);
            ExcelAddressBase a = Address;
            if (a._toRow <= a._fromRow) return; // no row below the buttons

            _ws.Cells[a._fromRow + 1, a._fromCol, a._toRow, a._toCol].Sort(column, descending);

            ClearSort();
            XmlDocument doc = Node.OwnerDocument;
            XmlElement sortState = doc.CreateElement("sortState", ExcelPackage.schemaMain);
            sortState.SetAttribute("ref", ExcelCellBase.GetAddress(a._fromRow + 1, a._fromCol, a._toRow, a._toCol));
            XmlElement condition = doc.CreateElement("sortCondition", ExcelPackage.schemaMain);
            if (descending) condition.SetAttribute("descending", "1");
            condition.SetAttribute("ref", ExcelCellBase.GetAddress(a._fromRow + 1, a._fromCol + column, a._toRow, a._fromCol + column));
            sortState.AppendChild(condition);

            if (_table == null)
            {
                // in a worksheet the sort of the autofilter is inside it, after the filter columns and before any extension list
                XmlNode ext = Node.SelectSingleNode("d:extLst", Ns);
                if (ext == null) Node.AppendChild(sortState);
                else Node.InsertBefore(sortState, ext);
            }
            else
            {
                // in a table it follows the autofilter
                Owner.InsertAfter(sortState, Node);
            }
            // the rows moved, so the hidden rows are no longer the right ones
            if (Node.SelectSingleNode("d:filterColumn", Ns) != null) _changed = true;
        }

        /// <summary>
        /// Removes the record of the sort. The rows stay where they are.
        /// </summary>
        public void ClearSort()
        {
            XmlNode old = _table == null ? Node.SelectSingleNode("d:sortState", Ns) : Owner.SelectSingleNode("d:sortState", Ns);
            if (old != null) old.ParentNode.RemoveChild(old);
        }

        /// <summary>
        /// Hides the rows that do not pass the filters and shows the ones that do.
        /// It is called when the package is saved if a filter was set or taken off since the last call.
        /// A kind of filter that is not handled here (top 10, dynamic, colour, icon, date groups) is left to Excel: its column lets every row pass, and no row is shown that was hidden.
        /// </summary>
        public void Apply()
        {
            _changed = false;
            ExcelAddressBase a = Address;
            var tests = new List<KeyValuePair<int, Func<int, bool>>>();
            bool unknown = false;
            foreach (XmlNode n in Node.SelectNodes("d:filterColumn", Ns))
            {
                var fc = (XmlElement)n;
                int colId;
                if (!int.TryParse(fc.GetAttribute("colId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out colId) || colId < 0 || a._fromCol + colId > a._toCol)
                {
                    unknown = true;
                    continue;
                }
                int col = a._fromCol + colId;
                XmlElement filters = fc.SelectSingleNode("d:filters", Ns) as XmlElement;
                XmlElement custom = fc.SelectSingleNode("d:customFilters", Ns) as XmlElement;
                if (filters != null && filters.SelectSingleNode("d:dateGroupItem", Ns) == null)
                {
                    tests.Add(new KeyValuePair<int, Func<int, bool>>(col, ValueTest(filters, col)));
                }
                else if (custom != null)
                {
                    tests.Add(new KeyValuePair<int, Func<int, bool>>(col, CustomTest(custom, col)));
                }
                else
                {
                    unknown = true;
                }
            }

            for (int row = a._fromRow + 1; row <= a._toRow; row++)
            {
                bool pass = true;
                foreach (var t in tests)
                {
                    if (!t.Value(row)) { pass = false; break; }
                }
                ExcelRow r = _ws.Row(row);
                if (!pass)
                {
                    if (!r.Hidden) r.Hidden = true;
                }
                else if (!unknown)
                {
                    if (r.Hidden) r.Hidden = false;
                }
            }
        }

        internal void ApplyIfChanged()
        {
            if (!_changed) return;
            if (Owner.SelectSingleNode("d:autoFilter", Ns) == null) { _changed = false; return; } // removed after it was changed
            Apply();
        }

        #region Writing the filter columns
        void CheckColumn(int column)
        {
            ExcelAddressBase a = Address;
            if (column < 0 || a._fromCol + column > a._toCol)
            {
                throw new ArgumentOutOfRangeException("column", "The autofilter " + a.Address + " has the columns 0 to " + (a._toCol - a._fromCol).ToString(CultureInfo.InvariantCulture));
            }
        }

        // A new filterColumn for the column, in place of the one it had, with an empty child of the given name.
        // The filter columns are kept in the order of their columns and come before sortState and extLst.
        XmlElement NewFilterColumn(int column, string childName)
        {
            XmlElement node = Node;
            XmlNode old = node.SelectSingleNode("d:filterColumn[@colId='" + column.ToString(CultureInfo.InvariantCulture) + "']", Ns);
            if (old != null) node.RemoveChild(old);

            XmlElement fc = node.OwnerDocument.CreateElement("filterColumn", ExcelPackage.schemaMain);
            fc.SetAttribute("colId", column.ToString(CultureInfo.InvariantCulture));
            XmlElement child = node.OwnerDocument.CreateElement(childName, ExcelPackage.schemaMain);
            fc.AppendChild(child);

            XmlNode before = null;
            foreach (XmlNode n in node.ChildNodes)
            {
                if (n.NodeType != XmlNodeType.Element) continue;
                if (n.LocalName == "filterColumn")
                {
                    int id;
                    if (int.TryParse(((XmlElement)n).GetAttribute("colId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out id) && id > column) { before = n; break; }
                }
                else
                {
                    before = n; // sortState or extLst
                    break;
                }
            }
            if (before == null) node.AppendChild(fc);
            else node.InsertBefore(fc, before);
            return child;
        }

        static void AddCustomFilter(XmlElement custom, eFilterOperator op, object value)
        {
            XmlElement f = custom.OwnerDocument.CreateElement("customFilter", ExcelPackage.schemaMain);
            if (op != eFilterOperator.Equal) f.SetAttribute("operator", s_operators[(int)op]); // equal is the default and is left out, as Excel does
            string text = ValueText(value);
            if (op == eFilterOperator.NotEqual && text.Length == 0) text = " "; // "not empty" is written with one space, as Excel does
            f.SetAttribute("val", text);
            custom.AppendChild(f);
        }

        // The value as it is written in the file: a number in the invariant form, a date as its serial number, a text as it is
        static string ValueText(object value)
        {
            if (value == null) return "";
            if (value is string) return (string)value;
            if (value is DateTime) return ((DateTime)value).ToOADate().ToString("R", CultureInfo.InvariantCulture);
            if (value is TimeSpan) return ((TimeSpan)value).TotalDays.ToString("R", CultureInfo.InvariantCulture);
            if (value is bool) return (bool)value ? "TRUE" : "FALSE";
            if (value is IFormattable) return ((IFormattable)value).ToString(null, CultureInfo.InvariantCulture);
            return value.ToString();
        }
        #endregion

        #region Testing a row
        string CellText(int row, int col)
        {
            return _ws.Cells[row, col].Text ?? "";
        }

        // The cell as a number: numbers, dates and times. False for texts, booleans and empty cells.
        bool CellNumber(int row, int col, out double number)
        {
            object v = _ws.GetValue(row, col);
            number = 0;
            if (v == null || v is string || v is bool) return false;
            if (v is DateTime) { number = ((DateTime)v).ToOADate(); return true; }
            if (v is TimeSpan) { number = ((TimeSpan)v).TotalDays; return true; }
            try
            {
                number = Convert.ToDouble(v, CultureInfo.InvariantCulture);
                return true;
            }
            catch (Exception)
            {
                return false; // a value that is not a number: an error value, an object
            }
        }

        Func<int, bool> ValueTest(XmlElement filters, int col)
        {
            string blankAttribute = filters.GetAttribute("blank");
            bool blank = blankAttribute == "1" || blankAttribute == "true";
            var values = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
            foreach (XmlNode f in filters.SelectNodes("d:filter", Ns)) values.Add(((XmlElement)f).GetAttribute("val"));
            return row =>
            {
                string text = CellText(row, col);
                return text.Length == 0 ? blank : values.Contains(text);
            };
        }

        Func<int, bool> CustomTest(XmlElement custom, int col)
        {
            string andAttribute = custom.GetAttribute("and");
            bool and = andAttribute == "1" || andAttribute == "true";
            var conditions = new List<Func<int, bool>>();
            foreach (XmlNode n in custom.SelectNodes("d:customFilter", Ns))
            {
                var f = (XmlElement)n;
                string opText = f.GetAttribute("operator");
                int op = opText.Length == 0 ? 0 : Array.IndexOf(s_operators, opText);
                if (op < 0) op = 0;
                conditions.Add(Condition((eFilterOperator)op, f.GetAttribute("val"), col));
            }
            return row =>
            {
                if (conditions.Count == 0) return true;
                foreach (var c in conditions)
                {
                    bool ok = c(row);
                    if (and && !ok) return false;
                    if (!and && ok) return true;
                }
                return and;
            };
        }

        // The rules below are what Excel 16 does when it applies the same conditions (filters-in-excel.ps1, and files in which
        // Excel set the conditions itself). They are not all what one would guess:
        //  - an empty value shows the empty cells, also with "not equal"; "not empty" is "not equal" to one space;
        //  - "equal" compares with the text the cell shows, so a 12 shown as 12.00 is not equal to 12;
        //  - "not equal" and the greater/less operators compare a number cell with a number value by the number;
        //  - a number cell does not match a text with wildcards, and is not greater or less than a text; a text cell is not
        //    greater or less than a number.
        Func<int, bool> Condition(eFilterOperator op, string val, int col)
        {
            if (val.Length == 0)
            {
                return row => CellText(row, col).Length == 0;
            }
            if (op == eFilterOperator.NotEqual && val.Trim().Length == 0)
            {
                return row => CellText(row, col).Length != 0;
            }

            double number;
            bool isNumber = double.TryParse(val, NumberStyles.Float, CultureInfo.InvariantCulture, out number);

            if (op == eFilterOperator.Equal || op == eFilterOperator.NotEqual)
            {
                var pattern = new Regex(WildcardPattern(val), RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Singleline);
                if (op == eFilterOperator.Equal)
                {
                    return row =>
                    {
                        double cell;
                        string text = CellText(row, col);
                        if (CellNumber(row, col, out cell)) return string.Equals(text, val, StringComparison.CurrentCultureIgnoreCase);
                        return pattern.IsMatch(text);
                    };
                }
                return row =>
                {
                    double cell;
                    if (CellNumber(row, col, out cell)) return !isNumber || cell != number;
                    return !pattern.IsMatch(CellText(row, col));
                };
            }

            if (isNumber)
            {
                return row =>
                {
                    double cell;
                    return CellNumber(row, col, out cell) && Compare(op, cell.CompareTo(number));
                };
            }

            // greater or less than a text: text cells only, in the order of the current culture without case
            return row =>
            {
                double cell;
                if (CellNumber(row, col, out cell)) return false;
                string text = CellText(row, col);
                if (text.Length == 0) return false;
                return Compare(op, string.Compare(text, val, StringComparison.CurrentCultureIgnoreCase));
            };
        }

        static bool Compare(eFilterOperator op, int order)
        {
            switch (op)
            {
                case eFilterOperator.Equal: return order == 0;
                case eFilterOperator.NotEqual: return order != 0;
                case eFilterOperator.GreaterThan: return order > 0;
                case eFilterOperator.GreaterThanOrEqual: return order >= 0;
                case eFilterOperator.LessThan: return order < 0;
                default: return order <= 0;
            }
        }

        // * is any run of characters, ? is one character, ~ before *, ? or ~ is that character itself
        static string WildcardPattern(string val)
        {
            var sb = new StringBuilder("^");
            for (int i = 0; i < val.Length; i++)
            {
                char ch = val[i];
                if (ch == '~' && i + 1 < val.Length && (val[i + 1] == '*' || val[i + 1] == '?' || val[i + 1] == '~'))
                {
                    sb.Append(Regex.Escape(val[i + 1].ToString()));
                    i++;
                }
                else if (ch == '*') sb.Append(".*");
                else if (ch == '?') sb.Append('.');
                else sb.Append(Regex.Escape(ch.ToString()));
            }
            sb.Append('$');
            return sb.ToString();
        }
        #endregion
    }
}
