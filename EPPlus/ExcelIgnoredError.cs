// orabooster-epplus: ignored errors of a worksheet (https://github.com/orascope/orabooster-epplus/issues/9).
// Written from ECMA-376 Part 1, 18.3.1 (the ignoredErrors and ignoredError elements of the worksheet part).
// (c) 2026 OraScope Co.,Ltd. GNU Lesser General Public License v3.0, as the rest of the library.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml
{
    /// <summary>
    /// The error checks of Excel (File, Options, Formulas, Error checking rules). A cell that fails a check gets a green triangle.
    /// The values can be combined.
    /// </summary>
    [Flags]
    public enum eErrorChecks
    {
        /// <summary>No check</summary>
        None = 0,
        /// <summary>The formula of the cell results in an error</summary>
        EvalError = 0x1,
        /// <summary>A date stored as text with a two-digit year</summary>
        TwoDigitTextYear = 0x2,
        /// <summary>A number stored as text, or a number that starts with an apostrophe</summary>
        NumberStoredAsText = 0x4,
        /// <summary>A formula that differs from the formulas around it</summary>
        Formula = 0x8,
        /// <summary>A formula that leaves out cells next to the range it refers to</summary>
        FormulaRange = 0x10,
        /// <summary>An unlocked cell with a formula</summary>
        UnlockedFormula = 0x20,
        /// <summary>A formula that refers to empty cells</summary>
        EmptyCellReference = 0x40,
        /// <summary>A value in a table that does not pass the data validation of its column</summary>
        ListDataValidation = 0x80,
        /// <summary>A cell of a calculated table column with a formula that differs from the column formula</summary>
        CalculatedColumn = 0x100,
        /// <summary>Every check</summary>
        All = 0x1FF
    }

    /// <summary>
    /// A range of a worksheet and the error checks Excel does not show for it.
    /// </summary>
    public class ExcelIgnoredError
    {
        // The attribute of the ignoredError element for each check, in the order of the schema
        static readonly KeyValuePair<eErrorChecks, string>[] s_attributes =
        {
            new KeyValuePair<eErrorChecks, string>(eErrorChecks.EvalError, "evalError"),
            new KeyValuePair<eErrorChecks, string>(eErrorChecks.TwoDigitTextYear, "twoDigitTextYear"),
            new KeyValuePair<eErrorChecks, string>(eErrorChecks.NumberStoredAsText, "numberStoredAsText"),
            new KeyValuePair<eErrorChecks, string>(eErrorChecks.Formula, "formula"),
            new KeyValuePair<eErrorChecks, string>(eErrorChecks.FormulaRange, "formulaRange"),
            new KeyValuePair<eErrorChecks, string>(eErrorChecks.UnlockedFormula, "unlockedFormula"),
            new KeyValuePair<eErrorChecks, string>(eErrorChecks.EmptyCellReference, "emptyCellReference"),
            new KeyValuePair<eErrorChecks, string>(eErrorChecks.ListDataValidation, "listDataValidation"),
            new KeyValuePair<eErrorChecks, string>(eErrorChecks.CalculatedColumn, "calculatedColumn"),
        };

        internal readonly XmlElement Node;

        internal ExcelIgnoredError(XmlElement node)
        {
            Node = node;
        }

        /// <summary>
        /// The cells, as one or more ranges separated by a space or a comma: "A2:A100", "A2:A100 C2:C100".
        /// Whole rows and columns are stored with their first and last cell.
        /// </summary>
        public string Address
        {
            get
            {
                return Node.GetAttribute("sqref");
            }
            set
            {
                SetAreas(ParseAreas(value));
            }
        }

        /// <summary>
        /// The checks that are not shown for the cells.
        /// </summary>
        public eErrorChecks Checks
        {
            get
            {
                eErrorChecks checks = eErrorChecks.None;
                foreach (var a in s_attributes)
                {
                    string v = Node.GetAttribute(a.Value);
                    if (v == "1" || v == "true") checks |= a.Key;
                }
                return checks;
            }
            set
            {
                if ((value & eErrorChecks.All) == 0) throw new ArgumentException("No check to ignore. To show every check again, remove the item from the collection.", "value");
                // sqref stays first, then the checks in the order of the schema
                string sqref = Node.GetAttribute("sqref");
                foreach (var a in s_attributes) Node.RemoveAttribute(a.Value);
                Node.SetAttribute("sqref", sqref);
                foreach (var a in s_attributes)
                {
                    if ((value & a.Key) != 0) Node.SetAttribute(a.Value, "1");
                }
            }
        }

        internal List<ExcelAddressBase> GetAreas()
        {
            return ParseAreas(Node.GetAttribute("sqref"));
        }

        internal void SetAreas(IEnumerable<ExcelAddressBase> areas)
        {
            Node.SetAttribute("sqref", ToSqref(areas));
        }

        /// <summary>
        /// The ranges as they are written in the file: separated by a space, each with its first and last cell.
        /// </summary>
        internal static string ToSqref(IEnumerable<ExcelAddressBase> areas)
        {
            var sb = new StringBuilder();
            foreach (var a in areas)
            {
                if (sb.Length > 0) sb.Append(' ');
                // corner by corner: GetAddress(fromRow, fromCol, toRow, toCol) shortens whole rows and columns to "D:D", Excel writes "D1:D1048576" here
                sb.Append(ExcelCellBase.GetAddress(a._fromRow, a._fromCol));
                if (a._toRow != a._fromRow || a._toCol != a._fromCol)
                {
                    sb.Append(':').Append(ExcelCellBase.GetAddress(a._toRow, a._toCol));
                }
            }
            return sb.ToString();
        }

        internal static List<ExcelAddressBase> ParseAreas(string address)
        {
            if (string.IsNullOrWhiteSpace(address)) throw new ArgumentException("The address is empty", "address");
            var areas = new List<ExcelAddressBase>();
            foreach (string part in address.Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (part.IndexOf('!') >= 0) throw new ArgumentException("The address must be on the worksheet itself, without a sheet name: " + part, "address");
                ExcelAddressBase a;
                try
                {
                    a = new ExcelAddressBase(part);
                }
                catch (Exception ex)
                {
                    // the address parser throws a plain Exception for text that is not an address
                    throw new ArgumentException("Not a cell range: " + part, "address", ex);
                }
                if (a._fromRow < 1 || a._fromCol < 1 || a._toRow > ExcelPackage.MaxRows || a._toCol > ExcelPackage.MaxColumns || a._fromRow > a._toRow || a._fromCol > a._toCol)
                {
                    throw new ArgumentException("Not a cell range: " + part, "address");
                }
                areas.Add(a);
            }
            return areas;
        }
    }

    /// <summary>
    /// The ranges of a worksheet for which Excel does not show given error checks.
    /// </summary>
    public class ExcelIgnoredErrorCollection : XmlHelper, IEnumerable<ExcelIgnoredError>
    {
        const string PATH = "d:ignoredErrors";
        readonly List<ExcelIgnoredError> _list = new List<ExcelIgnoredError>();
        readonly ExcelWorksheet _ws;

        // The elements that the schema (CT_Worksheet) puts before ignoredErrors. A new ignoredErrors element goes after the last of them.
        // CreateNode looks for the first later element it knows instead, and appends at the end when it finds none:
        // that is too late on a sheet whose later elements are not in its list (controls, drawingHF, mc:AlternateContent).
        static readonly HashSet<string> s_before = new HashSet<string>
        {
            "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges",
            "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting",
            "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
            "customProperties", "cellWatches"
        };

        internal ExcelIgnoredErrorCollection(XmlNamespaceManager nsm, XmlNode topNode, ExcelWorksheet ws)
            : base(nsm, topNode)
        {
            _ws = ws;
            SchemaNodeOrder = ws.SchemaNodeOrder;
            foreach (XmlNode n in topNode.SelectNodes(PATH + "/d:ignoredError", nsm))
            {
                var e = n as XmlElement;
                if (e != null) _list.Add(new ExcelIgnoredError(e));
            }
        }

        /// <summary>
        /// Adds a range for which the given checks are not shown.
        /// </summary>
        /// <param name="address">One or more ranges of this worksheet, separated by a space or a comma</param>
        /// <param name="checks">The checks to ignore</param>
        public ExcelIgnoredError Add(string address, eErrorChecks checks)
        {
            var areas = ExcelIgnoredError.ParseAreas(address);
            if ((checks & eErrorChecks.All) == 0) throw new ArgumentException("No check to ignore", "checks");

            XmlNode parent = GetOrCreateParent();
            XmlElement node = TopNode.OwnerDocument.CreateElement("ignoredError", ExcelPackage.schemaMain);
            parent.AppendChild(node);
            var item = new ExcelIgnoredError(node);
            item.SetAreas(areas);
            item.Checks = checks;
            _list.Add(item);
            return item;
        }

        /// <summary>
        /// Adds a range for which the given checks are not shown.
        /// </summary>
        /// <param name="address">A range of this worksheet, for example ws.Cells["A2:A100"]. It can hold several ranges.</param>
        /// <param name="checks">The checks to ignore</param>
        public ExcelIgnoredError Add(ExcelAddressBase address, eErrorChecks checks)
        {
            if (address == null) throw new ArgumentNullException("address");
            if (!string.IsNullOrEmpty(address.WorkSheet) && !address.WorkSheet.Equals(_ws.Name, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("The address is on another worksheet: " + address.WorkSheet, "address");
            }
            var areas = new List<ExcelAddressBase>();
            if (address.Addresses != null && address.Addresses.Count > 0) areas.AddRange(address.Addresses);
            else areas.Add(address);
            return Add(ExcelIgnoredError.ToSqref(areas), checks);
        }

        XmlNode GetOrCreateParent()
        {
            XmlNode parent = TopNode.SelectSingleNode(PATH, NameSpaceManager);
            if (parent != null) return parent;

            XmlNode last = null;
            foreach (XmlNode child in TopNode.ChildNodes)
            {
                if (child.NodeType == XmlNodeType.Element && child.NamespaceURI == ExcelPackage.schemaMain && s_before.Contains(child.LocalName)) last = child;
            }
            parent = TopNode.OwnerDocument.CreateElement("ignoredErrors", ExcelPackage.schemaMain);
            if (last == null) TopNode.PrependChild(parent);
            else TopNode.InsertAfter(parent, last);
            return parent;
        }

        /// <summary>
        /// Removes an item. Excel shows the checks for its cells again.
        /// </summary>
        public bool Remove(ExcelIgnoredError item)
        {
            if (item == null || !_list.Remove(item)) return false;
            XmlNode parent = item.Node.ParentNode;
            if (parent != null)
            {
                parent.RemoveChild(item.Node);
                // the schema does not allow an empty ignoredErrors element
                if (_list.Count == 0 && parent.ParentNode != null) parent.ParentNode.RemoveChild(parent);
            }
            return true;
        }

        /// <summary>
        /// Removes every item.
        /// </summary>
        public void Clear()
        {
            DeleteNode(PATH);
            _list.Clear();
        }

        /// <summary>The number of items</summary>
        public int Count
        {
            get { return _list.Count; }
        }

        /// <summary>The item at a position, in the order of the file</summary>
        public ExcelIgnoredError this[int index]
        {
            get { return _list[index]; }
        }

        /// <summary>
        /// Moves the ranges when rows or columns are inserted or deleted.
        /// </summary>
        /// <param name="from">The first row or column that is inserted or deleted</param>
        /// <param name="count">How many</param>
        /// <param name="rows">Rows (true) or columns (false)</param>
        /// <param name="insert">Inserted (true) or deleted (false)</param>
        internal void Shift(int from, int count, bool rows, bool insert)
        {
            for (int i = _list.Count - 1; i >= 0; i--)
            {
                ExcelIgnoredError item = _list[i];
                List<ExcelAddressBase> areas;
                try
                {
                    areas = item.GetAreas();
                }
                catch (ArgumentException)
                {
                    // an entry of a loaded file with an address that cannot be read here: left as it is, the insert or delete must not fail on it
                    continue;
                }
                var moved = new List<ExcelAddressBase>();
                foreach (ExcelAddressBase a in areas)
                {
                    ExcelAddressBase n = insert
                        ? (rows ? a.AddRow(from, count) : a.AddColumn(from, count))
                        : (rows ? a.DeleteRow(from, count) : a.DeleteColumn(from, count));
                    if (n == null) continue; // every cell of the range was deleted
                    // a range pushed past the last row or column of the sheet ends at the edge
                    int fromRow = n._fromRow, fromCol = n._fromCol;
                    int toRow = Math.Min(n._toRow, ExcelPackage.MaxRows), toCol = Math.Min(n._toCol, ExcelPackage.MaxColumns);
                    if (fromRow > toRow || fromCol > toCol) continue;
                    moved.Add(new ExcelAddressBase(fromRow, fromCol, toRow, toCol));
                }
                if (moved.Count == 0) Remove(item);
                else item.SetAreas(moved);
            }
        }

        /// <summary>The items, in the order of the file</summary>
        public IEnumerator<ExcelIgnoredError> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
    }
}
