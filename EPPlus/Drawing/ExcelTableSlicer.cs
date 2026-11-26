/*******************************************************************************
 * orabooster-epplus: this file is new in the fork (item B1 of issue #8, issue #48). It is not part of EPPlus 4.5 or EPPlusFree
 * and has no code from EPPlus 5 or later. Written by OraScope Co.,Ltd. 2026 from [MS-XLSX] (slicers: CT_Slicers, CT_Slicer,
 * CT_SlicerCacheDefinition, CT_TableSlicerCache, CT_SlicerRefs, CT_SlicerCaches), [MS-ODRAWXML] (the slicer graphic frame) and
 * from files Excel made, under the license of the library (LGPL-3.0).
 *
 * A slicer of a table: buttons with the values of a column of the table, which filter the table.
 *
 *   var slicer = sheet.Drawings.AddTableSlicer(table.Columns["Region"]);
 *   slicer.SetPosition(1, 0, 5, 0);
 *   slicer.Select("East", "West");      // the filter of the table column, which is what the slicer shows as selected
 *
 * Excel writes a table slicer in five places, and so does this class:
 *  - a slicer cache part (xl/slicerCaches/slicerCacheN.xml) with the table and the column, related to the workbook and listed
 *    in the extLst of the workbook (x15:slicerCaches), and a defined name with the name of the cache (#N/A);
 *  - a slicer part of the worksheet (xl/slicers/slicerN.xml) with a slicer element for each slicer of the sheet, related to the
 *    worksheet and listed in its extLst (x14:slicerList);
 *  - a drawing: an AlternateContent with the slicer graphic frame for Excel 2013 and later and a text box for older versions.
 * What the slicer shows as selected is the filter of the column in the autofilter of the table.
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security;
using System.Text;
using System.Xml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// A slicer of a table. Made with <see cref="ExcelDrawings.AddTableSlicer(ExcelTableColumn)"/>; a slicer of a table in a
    /// loaded workbook is in <see cref="ExcelWorksheet.Drawings"/> as one too. Excel 2013 and later show it; older versions show
    /// a text box in its place.
    /// </summary>
    public sealed class ExcelTableSlicer : ExcelDrawing
    {
        internal const string NsX14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
        internal const string NsX15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
        internal const string NsSle = "http://schemas.microsoft.com/office/drawing/2010/slicer";
        const string NsSle15 = "http://schemas.microsoft.com/office/drawing/2012/slicer";
        internal const string RelSlicer = "http://schemas.microsoft.com/office/2007/relationships/slicer";
        internal const string RelSlicerCache = "http://schemas.microsoft.com/office/2007/relationships/slicerCache";
        const string ContentTypeSlicer = "application/vnd.ms-excel.slicer+xml";
        const string ContentTypeSlicerCache = "application/vnd.ms-excel.slicerCache+xml";
        const string SheetExtUri = "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}";     // x14:slicerList of table slicers
        const string WorkbookExtUri = "{46BE6895-7355-4a93-B00E-2C351335B9C9}";  // x15:slicerCaches of table slicers
        const string CacheExtUri = "{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}";     // x15:tableSlicerCache
        const string NamePath = "mc:AlternateContent/mc:Choice/xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name";
        internal const string GraphicPath = "mc:AlternateContent/mc:Choice/xdr:graphicFrame/a:graphic/a:graphicData[@uri='" + NsSle + "']/sle:slicer";

        readonly SlicerPart _slicerPart;
        readonly XmlElement _slicer;
        readonly SlicerCache _cache;

        /// <summary>A new slicer of a table column.</summary>
        internal ExcelTableSlicer(ExcelDrawings drawings, XmlNode node, ExcelTableColumn column)
            : base(drawings, WithNamespaces(drawings, node), NamePath)
        {
            ExcelWorksheet ws = drawings.Worksheet;
            ExcelWorkbook wb = ws.Workbook;
            string slicerName = UniqueSlicerName(wb, column.Name);
            _cache = SlicerCache.Create(wb, column, UniqueCacheName(wb, column.Name));
            _slicerPart = SlicerPart.GetOrCreate(ws);
            _slicer = _slicerPart.Xml.CreateElement("slicer", NsX14);
            _slicer.SetAttribute("name", slicerName);
            _slicer.SetAttribute("cache", _cache.Name);
            _slicer.SetAttribute("caption", column.Name);
            _slicer.SetAttribute("rowHeight", "241300"); // required; what Excel writes for a new slicer (0.25 inch in EMU)
            _slicerPart.Xml.DocumentElement.AppendChild(_slicer);
            _slicerPart.Write();

            // the drawing: the slicer for Excel 2013 and later, a text box for older versions
            XmlDocument doc = node.OwnerDocument;
            // As Excel makes a slicer: absolute, so that the rows the slicer hides do not shrink it
            ((XmlElement)node).SetAttribute("editAs", "absolute");
            int frameId = 1; // an id no shape of the drawing part has
            foreach (XmlNode id in doc.SelectNodes("//xdr:cNvPr/@id", drawings.NameSpaceManager))
            {
                int n;
                if (int.TryParse(id.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out n) && n >= frameId) frameId = n + 1;
            }
            XmlElement alternate = doc.CreateElement("mc", "AlternateContent", ExcelPackage.schemaMarkupCompatibility);
            alternate.SetAttribute("xmlns:mc", ExcelPackage.schemaMarkupCompatibility);
            node.AppendChild(alternate);
            alternate.InnerXml =
                "<mc:Choice xmlns:sle15=\"" + NsSle15 + "\" Requires=\"sle15\">" +
                "<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"" + frameId.ToString(CultureInfo.InvariantCulture) + "\" name=\"" + Esc(slicerName) + "\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>" +
                "<xdr:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm>" +
                "<a:graphic><a:graphicData uri=\"" + NsSle + "\"><sle:slicer xmlns:sle=\"" + NsSle + "\" name=\"" + Esc(slicerName) + "\" /></a:graphicData></a:graphic>" +
                "</xdr:graphicFrame></mc:Choice>" +
                "<mc:Fallback><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"0\" name=\"\" /><xdr:cNvSpPr><a:spLocks noTextEdit=\"1\" /></xdr:cNvSpPr></xdr:nvSpPr>" +
                "<xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"1828800\" cy=\"2286000\" /></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst /></a:prstGeom>" +
                "<a:solidFill><a:prstClr val=\"white\" /></a:solidFill><a:ln w=\"1\"><a:solidFill><a:prstClr val=\"green\" /></a:solidFill></a:ln></xdr:spPr>" +
                "<xdr:txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\" /><a:lstStyle /><a:p><a:r><a:rPr lang=\"en-US\" sz=\"1100\" />" +
                "<a:t>This shape represents a table slicer. Table slicers are not supported in this version of Excel.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>";
            node.AppendChild(doc.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings));
            SetSize(192, 240); // what Excel makes: 2 by 2.5 inches
        }

        /// <summary>A slicer of a loaded workbook.</summary>
        ExcelTableSlicer(ExcelDrawings drawings, XmlNode node, SlicerPart slicerPart, XmlElement slicer, SlicerCache cache)
            : base(drawings, node, NamePath)
        {
            _slicerPart = slicerPart;
            _slicer = slicer;
            _cache = cache;
        }

        /// <summary>
        /// The slicer of a table in a drawing of a loaded workbook, or null when the drawing is not one (a slicer of a pivot table
        /// is not: it stays a plain drawing, as before).
        /// </summary>
        internal static ExcelTableSlicer TryLoad(ExcelDrawings drawings, XmlNode node)
        {
            WithNamespaces(drawings, node);
            XmlElement graphic = node.SelectSingleNode(GraphicPath, drawings.NameSpaceManager) as XmlElement;
            if (graphic == null) return null;
            string name = graphic.GetAttribute("name");
            ExcelWorksheet ws = drawings.Worksheet;
            foreach (SlicerPart part in SlicerPart.All(ws))
            {
                foreach (XmlElement slicer in part.Xml.DocumentElement.ChildNodes.OfElements("slicer", NsX14))
                {
                    if (slicer.GetAttribute("name") != name) continue;
                    SlicerCache cache = SlicerCache.Find(ws.Workbook, slicer.GetAttribute("cache"));
                    if (cache == null || !cache.IsTableCache) return null;
                    return new ExcelTableSlicer(drawings, node, part, slicer, cache);
                }
            }
            return null;
        }

        // The paths of this drawing use the prefixes mc and sle, which the drawings of a worksheet did not need before
        static XmlNode WithNamespaces(ExcelDrawings drawings, XmlNode node)
        {
            if (!drawings.NameSpaceManager.HasNamespace("mc")) drawings.NameSpaceManager.AddNamespace("mc", ExcelPackage.schemaMarkupCompatibility);
            if (!drawings.NameSpaceManager.HasNamespace("sle")) drawings.NameSpaceManager.AddNamespace("sle", NsSle);
            return node;
        }

        static string Esc(string text)
        {
            return SecurityElement.Escape(text ?? "");
        }

        /// <summary>The name of the slicer in the workbook (the slicer element); the drawing has it as its name too.</summary>
        public string SlicerName
        {
            get { return _slicer.GetAttribute("name"); }
        }

        /// <summary>The name of the slicer cache: the defined name Excel uses for the slicer in formulas (CUBE functions).</summary>
        public string CacheName
        {
            get { return _cache.Name; }
        }

        /// <summary>The table of the slicer, or null if the workbook has no table with the id of the cache.</summary>
        public ExcelTable Table
        {
            get { return _cache.FindTable(_drawings.Worksheet.Workbook); }
        }

        /// <summary>The column of the table that the slicer filters, or null if the table has no column with the id of the cache.</summary>
        public ExcelTableColumn TableColumn
        {
            get
            {
                ExcelTable table = Table;
                if (table == null) return null;
                foreach (ExcelTableColumn c in table.Columns)
                {
                    if (c.Id == _cache.ColumnId) return c;
                }
                return null;
            }
        }

        /// <summary>The text in the header of the slicer. The name of the column when the slicer is made.</summary>
        public string Caption
        {
            get { return _slicer.GetAttribute("caption"); }
            set { SetSlicerAttribute("caption", value); }
        }

        /// <summary>True if the slicer shows its header (the default).</summary>
        public bool ShowCaption
        {
            get { return _slicer.GetAttribute("showCaption") != "0"; }
            set { SetSlicerAttribute("showCaption", value ? null : "0"); }
        }

        /// <summary>The number of columns of buttons: 1 (the default) to 20000.</summary>
        public int ColumnCount
        {
            get
            {
                int count;
                return int.TryParse(_slicer.GetAttribute("columnCount"), NumberStyles.Integer, CultureInfo.InvariantCulture, out count) ? count : 1;
            }
            set
            {
                if (value < 1 || value > 20000) throw new ArgumentOutOfRangeException("value", "A slicer has 1 to 20000 columns of buttons");
                SetSlicerAttribute("columnCount", value == 1 ? null : value.ToString(CultureInfo.InvariantCulture));
            }
        }

        /// <summary>The height of a button, in EMU (914400 to an inch). 241300 when the slicer is made.</summary>
        public long RowHeight
        {
            get
            {
                long height;
                return long.TryParse(_slicer.GetAttribute("rowHeight"), NumberStyles.Integer, CultureInfo.InvariantCulture, out height) ? height : 241300;
            }
            set
            {
                if (value <= 0) throw new ArgumentOutOfRangeException("value", "The height of a button is more than 0");
                SetSlicerAttribute("rowHeight", value.ToString(CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// The name of the slicer style, as Excel lists them (SlicerStyleLight1 to SlicerStyleLight6, SlicerStyleOther1 and 2,
        /// SlicerStyleDark1 to SlicerStyleDark6), or null for the default of the workbook.
        /// </summary>
        public string Style
        {
            get
            {
                string style = _slicer.GetAttribute("style");
                return style.Length == 0 ? null : style;
            }
            set { SetSlicerAttribute("style", string.IsNullOrEmpty(value) ? null : value); }
        }

        /// <summary>
        /// The items that are selected: the values the filter of the table column shows. Empty when every item is (no filter).
        /// </summary>
        public string[] SelectedItems
        {
            get
            {
                ExcelTable table = Table;
                ExcelTableColumn column = TableColumn;
                if (table == null || column == null) return new string[0];
                XmlNode filters = table.TableXml.DocumentElement.SelectSingleNode(
                    "d:autoFilter/d:filterColumn[@colId='" + column.Position.ToString(CultureInfo.InvariantCulture) + "']/d:filters", _drawings.Worksheet.NameSpaceManager);
                var items = new List<string>();
                if (filters == null) return items.ToArray();
                foreach (XmlElement f in filters.ChildNodes.OfElements("filter", ExcelPackage.schemaMain)) items.Add(f.GetAttribute("val"));
                if (((XmlElement)filters).GetAttribute("blank") == "1") items.Add("");
                return items.ToArray();
            }
        }

        /// <summary>
        /// Selects these items and no other: the table shows the rows whose cell in the column has one of the texts, as the
        /// cell shows it (see <see cref="ExcelAutoFilter.ShowValues(int, string[])"/>; an empty text stands for the empty cells).
        /// The table must show its filter buttons (<see cref="ExcelTable.ShowFilter"/>).
        /// </summary>
        public void Select(params string[] items)
        {
            ExcelTable table = RequireTable();
            table.AutoFilter.ShowValues(TableColumn.Position, items);
        }

        /// <summary>Selects every item: the filter of the table column is taken off.</summary>
        public void ClearSelection()
        {
            ExcelTable table = RequireTable();
            if (table.ShowFilter) table.AutoFilter.Clear(TableColumn.Position);
        }

        ExcelTable RequireTable()
        {
            ExcelTable table = Table;
            if (table == null || TableColumn == null) throw new InvalidOperationException("The table or the column of the slicer is not in the workbook");
            return table;
        }

        void SetSlicerAttribute(string name, string value)
        {
            if (value == null) _slicer.RemoveAttribute(name);
            else _slicer.SetAttribute(name, value);
            _slicerPart.Write();
        }

        /// <summary>
        /// Removed from the worksheet: the drawing, the slicer element (with the slicer part when it was the last of the sheet),
        /// and the slicer cache with its defined name when no other slicer uses it.
        /// </summary>
        internal override void DeleteMe()
        {
            ExcelWorksheet ws = _drawings.Worksheet;
            _slicerPart.Xml.DocumentElement.RemoveChild(_slicer);
            if (_slicerPart.Xml.DocumentElement.ChildNodes.OfElements("slicer", NsX14).Count == 0) _slicerPart.Delete(ws);
            else _slicerPart.Write();
            bool used = false;
            foreach (ExcelWorksheet sheet in ws.Workbook.Worksheets)
            {
                if (sheet is ExcelChartsheet) continue;
                foreach (SlicerPart part in SlicerPart.All(sheet))
                {
                    foreach (XmlElement s in part.Xml.DocumentElement.ChildNodes.OfElements("slicer", NsX14))
                    {
                        if (s.GetAttribute("cache") == _cache.Name) used = true;
                    }
                }
            }
            if (!used) _cache.Delete(ws.Workbook);
            base.DeleteMe();
        }

        /// <summary>A slicer name that no slicer of the workbook has: the column name, then "name 1", "name 2" and so on.</summary>
        static string UniqueSlicerName(ExcelWorkbook wb, string name)
        {
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ExcelWorksheet ws in wb.Worksheets)
            {
                if (ws is ExcelChartsheet) continue;
                foreach (SlicerPart part in SlicerPart.All(ws))
                {
                    foreach (XmlElement s in part.Xml.DocumentElement.ChildNodes.OfElements("slicer", NsX14)) used.Add(s.GetAttribute("name"));
                }
                foreach (ExcelDrawing d in ws.Drawings) used.Add(d.Name);
            }
            if (!used.Contains(name)) return name;
            for (int i = 1; ; i++)
            {
                string candidate = name + " " + i.ToString(CultureInfo.InvariantCulture);
                if (!used.Contains(candidate)) return candidate;
            }
        }

        /// <summary>
        /// A cache name that is not a defined name of the workbook or the name of another cache: Slicer_ and the column name with
        /// what a defined name can not have made an underscore, then Slicer_name1, Slicer_name2 and so on.
        /// </summary>
        static string UniqueCacheName(ExcelWorkbook wb, string column)
        {
            var sb = new StringBuilder("Slicer_");
            foreach (char c in column) sb.Append(char.IsLetterOrDigit(c) || c == '_' || c == '.' ? c : '_');
            string name = sb.ToString();
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ExcelNamedRange n in wb.Names) used.Add(n.Name);
            foreach (string cacheName in SlicerCache.AllNames(wb)) used.Add(cacheName);
            if (!used.Contains(name)) return name;
            for (int i = 1; ; i++)
            {
                string candidate = name + i.ToString(CultureInfo.InvariantCulture);
                if (!used.Contains(candidate)) return candidate;
            }
        }

        internal static XmlNamespaceManager NewNs(XmlNameTable names)
        {
            var ns = new XmlNamespaceManager(names);
            ns.AddNamespace("d", ExcelPackage.schemaMain);
            ns.AddNamespace("x14", NsX14);
            ns.AddNamespace("x15", NsX15);
            ns.AddNamespace("r", ExcelPackage.schemaRelationships);
            return ns;
        }

        internal static void WriteXml(Packaging.ZipPackagePart part, XmlDocument xml)
        {
            // the stream belongs to the part and is read again (and when the package is written): it is not closed here
            xml.Save(part.GetStream(FileMode.Create, FileAccess.Write));
        }

        const string XmlDeclaration = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";

        static Uri NewUri(Packaging.ZipPackage package, string format)
        {
            for (int i = 1; ; i++)
            {
                var uri = new Uri(string.Format(CultureInfo.InvariantCulture, format, i), UriKind.Relative);
                if (!package.PartExists(uri)) return uri;
            }
        }

        /// <summary>The ext with this uri in the extLst under a top node, made at the end of the extLst when it is not there.</summary>
        static XmlElement GetOrCreateExt(XmlHelper owner, XmlNode top, string uri, string prefix, string ns)
        {
            XmlElement extLst = (XmlElement)owner.CreateNode("d:extLst");
            foreach (XmlElement ext in extLst.ChildNodes.OfElements("ext", ExcelPackage.schemaMain))
            {
                if (ext.GetAttribute("uri") == uri) return ext;
            }
            XmlElement newExt = top.OwnerDocument.CreateElement("ext", ExcelPackage.schemaMain);
            newExt.SetAttribute("uri", uri);
            newExt.SetAttribute("xmlns:" + prefix, ns);
            extLst.AppendChild(newExt);
            return newExt;
        }

        /// <summary>Takes a child out of an ext list (x14:slicer, x14:slicerCache by its r:id), and what is empty then.</summary>
        static void RemoveFromExt(XmlNode top, string uri, string relId)
        {
            XmlNode extLst = null;
            foreach (XmlNode child in top.ChildNodes)
            {
                if (child.LocalName == "extLst" && child.NamespaceURI == ExcelPackage.schemaMain) extLst = child;
            }
            if (extLst == null) return;
            foreach (XmlElement ext in extLst.ChildNodes.OfElements("ext", ExcelPackage.schemaMain))
            {
                if (ext.GetAttribute("uri") != uri) continue;
                foreach (XmlNode list in ext.ChildNodes)
                {
                    var remove = new List<XmlElement>();
                    foreach (XmlNode item in list.ChildNodes)
                    {
                        var e = item as XmlElement;
                        if (e != null && e.GetAttribute("id", ExcelPackage.schemaRelationships) == relId) remove.Add(e);
                    }
                    foreach (XmlElement e in remove) list.RemoveChild(e);
                }
                var empty = new List<XmlNode>();
                foreach (XmlNode list in ext.ChildNodes)
                {
                    if (list is XmlElement && list.ChildNodes.OfElements(null, null).Count == 0) empty.Add(list);
                }
                foreach (XmlNode list in empty) ext.RemoveChild(list);
                if (ext.ChildNodes.OfElements(null, null).Count == 0) extLst.RemoveChild(ext);
                break;
            }
            if (extLst.ChildNodes.OfElements(null, null).Count == 0) top.RemoveChild(extLst);
        }

        /// <summary>A slicer part of a worksheet (xl/slicers/slicerN.xml).</summary>
        internal sealed class SlicerPart
        {
            internal Uri Uri;
            internal Packaging.ZipPackagePart Part;
            internal Packaging.ZipPackageRelationship Relation;
            internal XmlDocument Xml;

            internal void Write()
            {
                WriteXml(Part, Xml);
            }

            /// <summary>The slicer parts of a worksheet, read from the package once and then kept with the worksheet.</summary>
            internal static List<SlicerPart> All(ExcelWorksheet ws)
            {
                if (ws._slicerParts != null) return ws._slicerParts;
                var list = new List<SlicerPart>();
                foreach (var rel in ws.Part.GetRelationshipsByType(RelSlicer))
                {
                    Uri uri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                    if (!ws._package.Package.PartExists(uri)) continue;
                    var part = ws._package.Package.GetPart(uri);
                    var xml = new XmlDocument();
                    XmlHelper.LoadXmlSafe(xml, part.GetStream());
                    list.Add(new SlicerPart { Uri = uri, Part = part, Relation = rel, Xml = xml });
                }
                ws._slicerParts = list;
                return list;
            }

            /// <summary>The slicer part of a worksheet for table slicers: the first one, or a new one.</summary>
            internal static SlicerPart GetOrCreate(ExcelWorksheet ws)
            {
                List<SlicerPart> all = All(ws);
                if (all.Count > 0) return all[0];
                var package = ws._package.Package;
                Uri uri = NewUri(package, "/xl/slicers/slicer{0}.xml");
                var part = package.CreatePart(uri, ContentTypeSlicer, ws._package.Compression);
                var rel = ws.Part.CreateRelationship(UriHelper.GetRelativeUri(ws.WorksheetUri, uri), Packaging.TargetMode.Internal, RelSlicer);
                var xml = new XmlDocument();
                xml.LoadXml(XmlDeclaration + "<slicers xmlns=\"" + NsX14 + "\" xmlns:mc=\"" + ExcelPackage.schemaMarkupCompatibility + "\" mc:Ignorable=\"x\" xmlns:x=\"" + ExcelPackage.schemaMain + "\" />");
                var sp = new SlicerPart { Uri = uri, Part = part, Relation = rel, Xml = xml };
                sp.Write();
                // listed in the extLst of the worksheet
                XmlElement ext = GetOrCreateExt(ws, ws.WorksheetXml.DocumentElement, SheetExtUri, "x15", NsX15);
                XmlElement list = ext.SelectSingleNode("x14:slicerList", NewNs(ws.WorksheetXml.NameTable)) as XmlElement;
                if (list == null)
                {
                    list = ws.WorksheetXml.CreateElement("x14", "slicerList", NsX14);
                    list.SetAttribute("xmlns:x14", NsX14);
                    ext.AppendChild(list);
                }
                XmlElement item = ws.WorksheetXml.CreateElement("x14", "slicer", NsX14);
                item.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);
                list.AppendChild(item);
                all.Add(sp);
                return sp;
            }

            /// <summary>Takes the part out of the package, with its relationship and its line in the extLst of the worksheet.</summary>
            internal void Delete(ExcelWorksheet ws)
            {
                RemoveFromExt(ws.WorksheetXml.DocumentElement, SheetExtUri, Relation.Id);
                ws.Part.DeleteRelationship(Relation.Id);
                ws._package.Package.DeletePart(Uri);
                All(ws).Remove(this);
            }
        }

        /// <summary>A slicer cache part (xl/slicerCaches/slicerCacheN.xml).</summary>
        internal sealed class SlicerCache
        {
            internal Uri Uri;
            internal Packaging.ZipPackagePart Part;
            internal Packaging.ZipPackageRelationship Relation;
            internal XmlDocument Xml;

            internal string Name { get { return Xml.DocumentElement.GetAttribute("name"); } }

            XmlElement TableCache
            {
                get { return Xml.DocumentElement.SelectSingleNode("x14:extLst/d:ext[@uri='" + CacheExtUri + "']/x15:tableSlicerCache", NewNs(Xml.NameTable)) as XmlElement; }
            }

            internal bool IsTableCache { get { return TableCache != null; } }

            internal int ColumnId
            {
                get
                {
                    int id;
                    return TableCache != null && int.TryParse(TableCache.GetAttribute("column"), NumberStyles.Integer, CultureInfo.InvariantCulture, out id) ? id : -1;
                }
            }

            internal ExcelTable FindTable(ExcelWorkbook wb)
            {
                int tableId;
                if (TableCache == null || !int.TryParse(TableCache.GetAttribute("tableId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out tableId)) return null;
                foreach (ExcelWorksheet ws in wb.Worksheets)
                {
                    if (ws is ExcelChartsheet) continue;
                    foreach (ExcelTable t in ws.Tables)
                    {
                        if (t.Id == tableId) return t;
                    }
                }
                return null;
            }

            /// <summary>Every slicer cache of the workbook (of tables and of pivot tables), read from the package.</summary>
            static IEnumerable<SlicerCache> All(ExcelWorkbook wb)
            {
                var package = wb._package.Package;
                foreach (var rel in wb.Part.GetRelationshipsByType(RelSlicerCache))
                {
                    Uri uri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                    if (!package.PartExists(uri)) continue;
                    var part = package.GetPart(uri);
                    var xml = new XmlDocument();
                    XmlHelper.LoadXmlSafe(xml, part.GetStream());
                    yield return new SlicerCache { Uri = uri, Part = part, Relation = rel, Xml = xml };
                }
            }

            internal static IEnumerable<string> AllNames(ExcelWorkbook wb)
            {
                foreach (SlicerCache c in All(wb)) yield return c.Name;
            }

            internal static SlicerCache Find(ExcelWorkbook wb, string name)
            {
                foreach (SlicerCache c in All(wb))
                {
                    if (c.Name == name) return c;
                }
                return null;
            }

            internal static SlicerCache Create(ExcelWorkbook wb, ExcelTableColumn column, string name)
            {
                var package = wb._package.Package;
                Uri uri = NewUri(package, "/xl/slicerCaches/slicerCache{0}.xml");
                var part = package.CreatePart(uri, ContentTypeSlicerCache, wb._package.Compression);
                var xml = new XmlDocument();
                xml.LoadXml(XmlDeclaration + "<slicerCacheDefinition xmlns=\"" + NsX14 + "\" xmlns:mc=\"" + ExcelPackage.schemaMarkupCompatibility + "\" mc:Ignorable=\"x\" xmlns:x=\"" + ExcelPackage.schemaMain + "\" " +
                    "name=\"" + Esc(name) + "\" sourceName=\"" + Esc(column.Name) + "\"><extLst><x:ext uri=\"" + CacheExtUri + "\" xmlns:x15=\"" + NsX15 + "\">" +
                    "<x15:tableSlicerCache tableId=\"" + column._tbl.Id.ToString(CultureInfo.InvariantCulture) + "\" column=\"" + column.Id.ToString(CultureInfo.InvariantCulture) + "\" />" +
                    "</x:ext></extLst></slicerCacheDefinition>");
                WriteXml(part, xml);
                var rel = wb.Part.CreateRelationship(UriHelper.GetRelativeUri(wb.WorkbookUri, uri), Packaging.TargetMode.Internal, RelSlicerCache);

                // listed in the extLst of the workbook
                XmlElement ext = GetOrCreateExt(wb, wb.WorkbookXml.DocumentElement, WorkbookExtUri, "x15", NsX15);
                XmlElement list = ext.SelectSingleNode("x15:slicerCaches", NewNs(wb.WorkbookXml.NameTable)) as XmlElement;
                if (list == null)
                {
                    list = wb.WorkbookXml.CreateElement("x15", "slicerCaches", NsX15);
                    list.SetAttribute("xmlns:x14", NsX14);
                    ext.AppendChild(list);
                }
                XmlElement item = wb.WorkbookXml.CreateElement("x14", "slicerCache", NsX14);
                item.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);
                list.AppendChild(item);

                // the defined name of the cache, as Excel writes it
                wb.Names.AddFormula(name, "#N/A");
                return new SlicerCache { Uri = uri, Part = part, Relation = rel, Xml = xml };
            }

            internal void Delete(ExcelWorkbook wb)
            {
                RemoveFromExt(wb.WorkbookXml.DocumentElement, WorkbookExtUri, Relation.Id);
                wb.Part.DeleteRelationship(Relation.Id);
                wb._package.Package.DeletePart(Uri);
                if (wb.Names.ContainsKey(Name)) wb.Names.Remove(Name);
            }
        }
    }

    internal static class XmlNodeListExtensions
    {
        /// <summary>The element children with this name and namespace (any element when the name is null).</summary>
        internal static List<XmlElement> OfElements(this XmlNodeList nodes, string localName, string ns)
        {
            var list = new List<XmlElement>();
            foreach (XmlNode n in nodes)
            {
                var e = n as XmlElement;
                if (e == null) continue;
                if (localName == null || (e.LocalName == localName && e.NamespaceURI == ns)) list.Add(e);
            }
            return list;
        }
    }
}
