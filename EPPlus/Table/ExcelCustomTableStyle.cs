/*******************************************************************************
 * orabooster-epplus: this file is new in the fork (item B3 of issue #8). It is not part of EPPlus 4.5 or EPPlusFree and has
 * no code from EPPlus 5 or later. Written by OraScope Co.,Ltd. 2026 from ECMA-376 Part 1 (18.8.39 tableStyle, 18.8.40
 * tableStyleElement, 18.8.42 tableStyles, 18.18.83 ST_TableStyleType), under the license of the library (LGPL-3.0).
 *
 * Table styles that are defined in the workbook ("custom table styles").
 *
 *   var style = package.Workbook.Styles.TableStyles.Add("Ledger");
 *   style.HeaderRow.Font.Bold = true;
 *   style.HeaderRow.Fill.PatternType = ExcelFillStyle.Solid;
 *   style.HeaderRow.Fill.BackgroundColor.Color = Color.FromArgb(0x1F, 0x4E, 0x78);
 *   style.FirstRowStripe.Fill.PatternType = ExcelFillStyle.Solid;
 *   style.FirstRowStripe.Fill.BackgroundColor.Color = Color.FromArgb(0xDD, 0xEB, 0xF7);
 *   table.StyleName = "Ledger";
 *
 * In the file a table style is a tableStyle element of styles.xml with one tableStyleElement for each part of a table that
 * has a format; the format is a dxf (a differential format, as conditional formatting uses) in the dxfs list.
 * A style that was loaded keeps its XML as it is unless it was changed here.
 *******************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Style.Dxf;

namespace OfficeOpenXml.Table
{
    /// <summary>
    /// The parts of a table that a custom table style can format.
    /// </summary>
    public enum eTableStyleElement
    {
        /// <summary>The whole table: every cell, where no other part says otherwise</summary>
        WholeTable,
        /// <summary>The header row</summary>
        HeaderRow,
        /// <summary>The total row</summary>
        TotalRow,
        /// <summary>The first column (shown when ShowFirstColumn of the table is on)</summary>
        FirstColumn,
        /// <summary>The last column (shown when ShowLastColumn of the table is on)</summary>
        LastColumn,
        /// <summary>The first band of rows (shown when ShowRowStripes of the table is on)</summary>
        FirstRowStripe,
        /// <summary>The second band of rows</summary>
        SecondRowStripe,
        /// <summary>The first band of columns (shown when ShowColumnStripes of the table is on)</summary>
        FirstColumnStripe,
        /// <summary>The second band of columns</summary>
        SecondColumnStripe,
        /// <summary>The first cell of the header row</summary>
        FirstHeaderCell,
        /// <summary>The last cell of the header row</summary>
        LastHeaderCell,
        /// <summary>The first cell of the total row</summary>
        FirstTotalCell,
        /// <summary>The last cell of the total row</summary>
        LastTotalCell
    }

    /// <summary>
    /// A table style that is defined in the workbook. Get one from <see cref="ExcelCustomTableStyleCollection.Add(string)"/> and
    /// give its name to <see cref="ExcelTable.StyleName"/>.
    /// </summary>
    public sealed class ExcelCustomTableStyle
    {
        /// <summary>One tableStyleElement: a part of the table, its format, and for a band its size.</summary>
        internal sealed class Part
        {
            internal string Type;                               // the name in the file (ST_TableStyleType), also for parts not in eTableStyleElement
            internal int Size = 1;                              // the rows or columns of a band
            internal ExcelDxfStyleConditionalFormatting Dxf;    // the format; null for a loaded part whose dxf is not there
            internal int SavedDxfId = -1;                       // its place in the dxfs list of the file, -1 when not written yet
            internal string SavedId;                            // Dxf.Id when it was loaded or written: another value means it was changed
            internal int SavedSize = 1;
        }

        readonly ExcelStyles _styles;
        readonly List<Part> _parts = new List<Part>();
        internal XmlElement Node;       // the tableStyle element; null until the style is written
        internal bool Written;          // false: made here and not written yet

        internal ExcelCustomTableStyle(ExcelStyles styles, string name)
        {
            _styles = styles;
            Name = name;
        }

        /// <summary>A style of a loaded workbook.</summary>
        internal ExcelCustomTableStyle(ExcelStyles styles, XmlElement node)
        {
            _styles = styles;
            Node = node;
            Written = true;
            Name = node.GetAttribute("name");
            foreach (XmlNode child in node.ChildNodes)
            {
                XmlElement element = child as XmlElement;
                if (element == null || element.LocalName != "tableStyleElement") continue;
                var part = new Part { Type = element.GetAttribute("type") };
                int size;
                if (int.TryParse(element.GetAttribute("size"), NumberStyles.Integer, CultureInfo.InvariantCulture, out size) && size > 0) part.Size = size;
                part.SavedSize = part.Size;
                int dxfId;
                if (int.TryParse(element.GetAttribute("dxfId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out dxfId) && dxfId >= 0 && dxfId < styles.Dxfs.Count)
                {
                    // a copy: a conditional format may use the same dxf, and a change made here must not reach it
                    part.Dxf = styles.Dxfs[dxfId].Clone();
                    part.SavedDxfId = dxfId;
                    part.SavedId = part.Dxf.Id;
                }
                _parts.Add(part);
            }
        }

        /// <summary>The name of the style, as in <see cref="ExcelTable.StyleName"/>.</summary>
        public string Name { get; private set; }

        static string TypeName(eTableStyleElement element)
        {
            string s = element.ToString();
            return char.ToLowerInvariant(s[0]) + s.Substring(1); // WholeTable -> wholeTable, as ST_TableStyleType
        }

        Part GetPart(eTableStyleElement element, bool create)
        {
            string type = TypeName(element);
            foreach (Part p in _parts)
            {
                if (p.Type == type)
                {
                    if (p.Dxf == null && create) p.Dxf = _styles.NewDxf();
                    return p;
                }
            }
            if (!create) return null;
            var part = new Part { Type = type, Dxf = _styles.NewDxf() };
            _parts.Add(part);
            return part;
        }

        /// <summary>
        /// The format of a part of the table: font, fill and border. A part with no format set is not written.
        /// </summary>
        /// <param name="element">The part of the table</param>
        public ExcelDxfStyleConditionalFormatting this[eTableStyleElement element]
        {
            get { return GetPart(element, true).Dxf; }
        }

        /// <summary>The format of the whole table: every cell, where no other part says otherwise.</summary>
        public ExcelDxfStyleConditionalFormatting WholeTable { get { return this[eTableStyleElement.WholeTable]; } }
        /// <summary>The format of the header row.</summary>
        public ExcelDxfStyleConditionalFormatting HeaderRow { get { return this[eTableStyleElement.HeaderRow]; } }
        /// <summary>The format of the total row.</summary>
        public ExcelDxfStyleConditionalFormatting TotalRow { get { return this[eTableStyleElement.TotalRow]; } }
        /// <summary>The format of the first column. Shown when <see cref="ExcelTable.ShowFirstColumn"/> is on.</summary>
        public ExcelDxfStyleConditionalFormatting FirstColumn { get { return this[eTableStyleElement.FirstColumn]; } }
        /// <summary>The format of the last column. Shown when <see cref="ExcelTable.ShowLastColumn"/> is on.</summary>
        public ExcelDxfStyleConditionalFormatting LastColumn { get { return this[eTableStyleElement.LastColumn]; } }
        /// <summary>The format of the first band of rows. Shown when <see cref="ExcelTable.ShowRowStripes"/> is on.</summary>
        public ExcelDxfStyleConditionalFormatting FirstRowStripe { get { return this[eTableStyleElement.FirstRowStripe]; } }
        /// <summary>The format of the second band of rows.</summary>
        public ExcelDxfStyleConditionalFormatting SecondRowStripe { get { return this[eTableStyleElement.SecondRowStripe]; } }
        /// <summary>The format of the first band of columns. Shown when <see cref="ExcelTable.ShowColumnStripes"/> is on.</summary>
        public ExcelDxfStyleConditionalFormatting FirstColumnStripe { get { return this[eTableStyleElement.FirstColumnStripe]; } }
        /// <summary>The format of the second band of columns.</summary>
        public ExcelDxfStyleConditionalFormatting SecondColumnStripe { get { return this[eTableStyleElement.SecondColumnStripe]; } }
        /// <summary>The format of the first cell of the header row.</summary>
        public ExcelDxfStyleConditionalFormatting FirstHeaderCell { get { return this[eTableStyleElement.FirstHeaderCell]; } }
        /// <summary>The format of the last cell of the header row.</summary>
        public ExcelDxfStyleConditionalFormatting LastHeaderCell { get { return this[eTableStyleElement.LastHeaderCell]; } }
        /// <summary>The format of the first cell of the total row.</summary>
        public ExcelDxfStyleConditionalFormatting FirstTotalCell { get { return this[eTableStyleElement.FirstTotalCell]; } }
        /// <summary>The format of the last cell of the total row.</summary>
        public ExcelDxfStyleConditionalFormatting LastTotalCell { get { return this[eTableStyleElement.LastTotalCell]; } }

        /// <summary>
        /// True if the style has a format for this part of the table: one that was set here, or one of the file (also when it
        /// only has what there is no property for, such as the borders between the cells).
        /// </summary>
        public bool HasElement(eTableStyleElement element)
        {
            Part p = GetPart(element, false);
            return p != null && p.Dxf != null && (p.SavedDxfId >= 0 || p.Dxf.HasValue);
        }

        /// <summary>
        /// The number of rows or columns in a band: 1 to 9, for <see cref="eTableStyleElement.FirstRowStripe"/>,
        /// <see cref="eTableStyleElement.SecondRowStripe"/>, <see cref="eTableStyleElement.FirstColumnStripe"/> and
        /// <see cref="eTableStyleElement.SecondColumnStripe"/>. 1 when it was not set.
        /// </summary>
        public int GetBandSize(eTableStyleElement stripe)
        {
            CheckStripe(stripe);
            Part p = GetPart(stripe, false);
            return p == null ? 1 : p.Size;
        }

        /// <summary>Sets the number of rows or columns in a band, 1 to 9. See <see cref="GetBandSize(eTableStyleElement)"/>.</summary>
        public void SetBandSize(eTableStyleElement stripe, int size)
        {
            CheckStripe(stripe);
            if (size < 1 || size > 9) throw new ArgumentOutOfRangeException("size", "The size of a band is 1 to 9");
            GetPart(stripe, true).Size = size;
        }

        static void CheckStripe(eTableStyleElement stripe)
        {
            if (stripe != eTableStyleElement.FirstRowStripe && stripe != eTableStyleElement.SecondRowStripe
                && stripe != eTableStyleElement.FirstColumnStripe && stripe != eTableStyleElement.SecondColumnStripe)
            {
                throw new ArgumentException("Only the row and column stripes have a band size", "stripe");
            }
        }

        /// <summary>True if the style must be written: it is new, or a part was changed since it was loaded or written.</summary>
        internal bool IsChanged
        {
            get
            {
                if (!Written) return true;
                foreach (Part p in _parts)
                {
                    if (p.Size != p.SavedSize) return true;
                    if (p.Dxf == null) continue;
                    if (p.SavedDxfId < 0) { if (p.Dxf.HasValue) return true; continue; } // a part that was added
                    if (p.Dxf.Id != p.SavedId) return true;
                }
                return false;
            }
        }

        /// <summary>
        /// Writes the style: its tableStyle element under tableStyles, and a dxf for every part that is new or was changed.
        /// A part that was not changed keeps its dxf.
        /// </summary>
        internal void Save(XmlElement tableStylesNode, XmlElement dxfsNode)
        {
            XmlDocument doc = tableStylesNode.OwnerDocument;
            if (Node == null)
            {
                Node = doc.CreateElement("tableStyle", ExcelPackage.schemaMain);
                Node.SetAttribute("name", Name);
                Node.SetAttribute("pivot", "0"); // a style for tables: Excel does not list it for pivot tables
                tableStylesNode.AppendChild(Node);
            }
            while (Node.FirstChild != null) Node.RemoveChild(Node.FirstChild);
            int count = 0;
            foreach (Part p in _parts)
            {
                // a part is written when it has a format, or when it is a band whose size was set (a band with no format of its
                // own still takes rows or columns: the dxfId of a tableStyleElement is optional)
                bool hasFormat = p.Dxf != null && (p.SavedDxfId >= 0 || p.Dxf.HasValue);
                if (!hasFormat && p.Size == 1) continue;           // looked at, but nothing was set
                if (hasFormat && (p.SavedDxfId < 0 || p.Dxf.Id != p.SavedId))
                {
                    // a new dxf at the end of the list, a copy of the format as it is now. A dxf that is in the file is never
                    // changed in place: a conditional format or another style may use it too
                    // The copy of a format that was loaded starts from its XML, which keeps what this library has no property for.
                    XmlNode basedOn = null;
                    if (p.SavedDxfId >= 0)
                    {
                        int n = 0;
                        foreach (XmlNode child in dxfsNode.ChildNodes)
                        {
                            if (!(child is XmlElement)) continue;
                            if (n++ == p.SavedDxfId) { basedOn = child; break; }
                        }
                    }
                    p.SavedDxfId = _styles.AppendDxf(dxfsNode, p.Dxf.Clone(), basedOn);
                    p.SavedId = p.Dxf.Id;
                }
                XmlElement element = doc.CreateElement("tableStyleElement", ExcelPackage.schemaMain);
                element.SetAttribute("type", p.Type);
                if (p.Size != 1) element.SetAttribute("size", p.Size.ToString(CultureInfo.InvariantCulture));
                if (hasFormat) element.SetAttribute("dxfId", p.SavedDxfId.ToString(CultureInfo.InvariantCulture));
                Node.AppendChild(element);
                p.SavedSize = p.Size;
                count++;
            }
            Node.SetAttribute("count", count.ToString(CultureInfo.InvariantCulture));
            Written = true;
        }
    }

    /// <summary>
    /// The table styles that are defined in the workbook: <see cref="ExcelStyles.TableStyles"/>.
    /// The styles Excel has built in (TableStyleLight1 and so on) are not in this list.
    /// </summary>
    public sealed class ExcelCustomTableStyleCollection : IEnumerable<ExcelCustomTableStyle>
    {
        readonly ExcelStyles _styles;
        readonly List<ExcelCustomTableStyle> _list = new List<ExcelCustomTableStyle>();
        bool _removed;

        internal ExcelCustomTableStyleCollection(ExcelStyles styles, XmlNode tableStylesNode)
        {
            _styles = styles;
            if (tableStylesNode == null) return;
            foreach (XmlNode child in tableStylesNode.ChildNodes)
            {
                XmlElement element = child as XmlElement;
                if (element != null && element.LocalName == "tableStyle") _list.Add(new ExcelCustomTableStyle(styles, element));
            }
        }

        /// <summary>The number of styles.</summary>
        public int Count { get { return _list.Count; } }

        /// <summary>A style by its position.</summary>
        public ExcelCustomTableStyle this[int index] { get { return _list[index]; } }

        /// <summary>A style by its name (case does not matter), or null.</summary>
        public ExcelCustomTableStyle this[string name]
        {
            get
            {
                foreach (ExcelCustomTableStyle s in _list)
                {
                    if (string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase)) return s;
                }
                return null;
            }
        }

        /// <summary>True if the workbook has a style with this name.</summary>
        public bool Contains(string name) { return this[name] != null; }

        /// <summary>
        /// Adds a table style with no formats. Set the formats of its parts, and give the name to
        /// <see cref="ExcelTable.StyleName"/> of the tables that use it.
        /// </summary>
        /// <param name="name">
        /// The name: not empty, at most 255 characters, not the name of a style of the workbook, and not one of the names
        /// Excel uses for its own styles (they start with TableStyle or PivotStyle).
        /// </param>
        public ExcelCustomTableStyle Add(string name)
        {
            if (string.IsNullOrEmpty(name) || name.Trim().Length == 0) throw new ArgumentException("The name of a table style can not be empty", "name");
            if (name.Length > 255) throw new ArgumentException("The name of a table style has at most 255 characters", "name");
            if (name.StartsWith("TableStyle", StringComparison.OrdinalIgnoreCase) || name.StartsWith("PivotStyle", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("Names that start with TableStyle or PivotStyle are the names of the styles of Excel", "name");
            }
            if (Contains(name)) throw new ArgumentException("The workbook already has a table style named " + name, "name");
            var style = new ExcelCustomTableStyle(_styles, name);
            _list.Add(style);
            return style;
        }

        /// <summary>
        /// Takes a style out of the workbook. Tables that use its name keep the name; Excel shows them with no style.
        /// </summary>
        /// <returns>False if there is no style with this name</returns>
        public bool Remove(string name)
        {
            ExcelCustomTableStyle style = this[name];
            if (style == null) return false;
            _list.Remove(style);
            if (style.Node != null && style.Node.ParentNode != null)
            {
                style.Node.ParentNode.RemoveChild(style.Node);
                _removed = true;
            }
            return true;
        }

        /// <summary>True if anything must be written.</summary>
        internal bool IsChanged
        {
            get
            {
                if (_removed) return true;
                foreach (ExcelCustomTableStyle s in _list)
                {
                    if (s.IsChanged) return true;
                }
                return false;
            }
        }

        /// <summary>Writes the styles that are new or were changed, and the number of styles.</summary>
        internal void Save(XmlElement tableStylesNode, XmlElement dxfsNode)
        {
            foreach (ExcelCustomTableStyle s in _list)
            {
                if (s.IsChanged) s.Save(tableStylesNode, dxfsNode);
            }
            int count = 0;
            foreach (XmlNode child in tableStylesNode.ChildNodes)
            {
                if (child is XmlElement && child.LocalName == "tableStyle") count++;
            }
            tableStylesNode.SetAttribute("count", count.ToString(CultureInfo.InvariantCulture));
            _removed = false;
        }

        /// <summary>The styles, in the order of the workbook.</summary>
        public IEnumerator<ExcelCustomTableStyle> GetEnumerator() { return _list.GetEnumerator(); }
        IEnumerator IEnumerable.GetEnumerator() { return _list.GetEnumerator(); }
    }
}
