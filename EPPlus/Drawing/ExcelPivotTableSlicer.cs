/*******************************************************************************
 * orabooster-epplus: this file is new in the fork (item B2 of issue #8, issue #65). It is not part of EPPlus 4.5 or EPPlusFree
 * and has no code from EPPlus 5 or later. Written by OraScope Co.,Ltd. 2026 from [MS-XLSX] (CT_SlicerCacheDefinition,
 * CT_SlicerCachePivotTables, CT_TabularSlicerCache, CT_TabularSlicerCacheItems, CT_PivotCacheDefinition of x14),
 * [MS-ODRAWXML] (the slicer graphic frame) and from files Excel made, under the license of the library (LGPL-3.0).
 *
 * A slicer of a pivot table: buttons with the items of a field, which filter the pivot table.
 *
 *   var slicer = sheet.Drawings.AddPivotTableSlicer(pivot.Fields["Region"]);
 *   slicer.SetPosition(1, 0, 6, 0);
 *   slicer.Select("East", "West");      // the items of the pivot field that are shown
 *
 * It is written as a table slicer (ExcelTableSlicer), with these differences, all as Excel has them: the slicer cache
 * names the pivot table (the id of its sheet and its name) and has the items of the field with what is selected; the pivot
 * cache gets an id (x14:pivotCacheDefinition) that the slicer cache names; the workbook and the worksheet list the parts in
 * the ext elements of x14; the drawing asks for a14. What is selected is the items of the pivot field that are not hidden.
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// A slicer of a pivot table. Made with <see cref="ExcelDrawings.AddPivotTableSlicer(ExcelPivotTableField)"/>; a slicer of
    /// a pivot table in a loaded workbook is in <see cref="ExcelWorksheet.Drawings"/> as one too. Excel 2010 and later show
    /// it; older versions show a text box in its place.
    /// </summary>
    public sealed class ExcelPivotTableSlicer : ExcelDrawing
    {
        const string NsA14 = "http://schemas.microsoft.com/office/drawing/2010/main";
        const string PivotCacheExtUri = "{725AE2AE-9491-48be-B2B4-4EB974FC3084}"; // x14:pivotCacheDefinition of a pivot cache

        readonly ExcelTableSlicer.SlicerPart _slicerPart;
        readonly XmlElement _slicer;
        readonly ExcelTableSlicer.SlicerCache _cache;
        ExcelPivotTableField _field;

        /// <summary>A new slicer of a field of a pivot table.</summary>
        internal ExcelPivotTableSlicer(ExcelDrawings drawings, XmlNode node, ExcelPivotTableField field)
            : base(drawings, WithNamespaces(drawings, node), ExcelTableSlicer.NamePath)
        {
            ExcelWorksheet ws = drawings.Worksheet;
            ExcelWorkbook wb = ws.Workbook;
            ExcelPivotTable pivot = field._table;
            _field = field;
            field.ItemsForSlicer(); // the field gets its items; throws for a field that can have none

            string slicerName = ExcelTableSlicer.UniqueSlicerName(wb, field.Name);
            string cacheName = ExcelTableSlicer.UniqueCacheName(wb, field.Name);
            string cacheId = PivotCacheId(pivot);
            MarkAsExcel2010(pivot);

            // the slicer cache: the pivot table and the items
            var package = wb._package.Package;
            Uri uri = ExcelTableSlicer.NewUri(package, "/xl/slicerCaches/slicerCache{0}.xml");
            var part = package.CreatePart(uri, ExcelTableSlicer.ContentTypeSlicerCache, wb._package.Compression);
            var xml = new XmlDocument();
            xml.LoadXml(ExcelTableSlicer.XmlDeclaration + "<slicerCacheDefinition xmlns=\"" + ExcelTableSlicer.NsX14 + "\" xmlns:mc=\"" + ExcelPackage.schemaMarkupCompatibility + "\" mc:Ignorable=\"x\" xmlns:x=\"" + ExcelPackage.schemaMain + "\" " +
                "name=\"" + ExcelTableSlicer.Esc(cacheName) + "\" sourceName=\"" + ExcelTableSlicer.Esc(field.Name) + "\">" +
                "<pivotTables><pivotTable tabId=\"" + pivot.WorkSheet.SheetID.ToString(CultureInfo.InvariantCulture) + "\" name=\"" + ExcelTableSlicer.Esc(pivot.Name) + "\" /></pivotTables>" +
                "<data><tabular pivotCacheId=\"" + cacheId + "\"><items count=\"0\" /></tabular></data></slicerCacheDefinition>");
            var rel = wb.Part.CreateRelationship(UriHelper.GetRelativeUri(wb.WorkbookUri, uri), Packaging.TargetMode.Internal, ExcelTableSlicer.RelSlicerCache);
            _cache = new ExcelTableSlicer.SlicerCache { Uri = uri, Part = part, Relation = rel, Xml = xml };
            // listed in the extLst of the workbook, in the ext of x14, which Excel has before the one of the table slicers
            XmlElement ext = ExcelTableSlicer.GetOrCreateExt(wb, wb.WorkbookXml.DocumentElement, ExcelTableSlicer.PivotWorkbookExtUri, "x14", ExcelTableSlicer.NsX14, ExcelTableSlicer.WorkbookExtUri);
            XmlElement list = ext.SelectSingleNode("x14:slicerCaches", ExcelTableSlicer.NewNs(wb.WorkbookXml.NameTable)) as XmlElement;
            if (list == null)
            {
                list = wb.WorkbookXml.CreateElement("x14", "slicerCaches", ExcelTableSlicer.NsX14);
                ext.AppendChild(list);
            }
            XmlElement item = wb.WorkbookXml.CreateElement("x14", "slicerCache", ExcelTableSlicer.NsX14);
            item.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);
            list.AppendChild(item);
            wb.Names.AddFormula(cacheName, "#N/A");

            // the slicer element, in the slicer part of the sheet for pivot table slicers
            _slicerPart = ExcelTableSlicer.SlicerPart.GetOrCreate(ws, true);
            _slicer = _slicerPart.Xml.CreateElement("slicer", ExcelTableSlicer.NsX14);
            _slicer.SetAttribute("name", slicerName);
            _slicer.SetAttribute("cache", cacheName);
            _slicer.SetAttribute("caption", field.Name);
            _slicer.SetAttribute("rowHeight", "241300");
            _slicerPart.Xml.DocumentElement.AppendChild(_slicer);
            _slicerPart.Write();

            // the drawing: the slicer for Excel 2010 and later, a text box for older versions
            XmlDocument doc = node.OwnerDocument;
            ((XmlElement)node).SetAttribute("editAs", "oneCell");
            int frameId = 1;
            foreach (XmlNode id in doc.SelectNodes("//xdr:cNvPr/@id", drawings.NameSpaceManager))
            {
                int n;
                if (int.TryParse(id.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out n) && n >= frameId) frameId = n + 1;
            }
            XmlElement alternate = doc.CreateElement("mc", "AlternateContent", ExcelPackage.schemaMarkupCompatibility);
            alternate.SetAttribute("xmlns:mc", ExcelPackage.schemaMarkupCompatibility);
            node.AppendChild(alternate);
            string esc = ExcelTableSlicer.Esc(slicerName);
            alternate.InnerXml =
                "<mc:Choice xmlns:a14=\"" + NsA14 + "\" Requires=\"a14\">" +
                "<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"" + frameId.ToString(CultureInfo.InvariantCulture) + "\" name=\"" + esc + "\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>" +
                "<xdr:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm>" +
                "<a:graphic><a:graphicData uri=\"" + ExcelTableSlicer.NsSle + "\"><sle:slicer xmlns:sle=\"" + ExcelTableSlicer.NsSle + "\" name=\"" + esc + "\" /></a:graphicData></a:graphic>" +
                "</xdr:graphicFrame></mc:Choice>" +
                "<mc:Fallback><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"0\" name=\"\" /><xdr:cNvSpPr><a:spLocks noTextEdit=\"1\" /></xdr:cNvSpPr></xdr:nvSpPr>" +
                "<xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"1828800\" cy=\"2286000\" /></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst /></a:prstGeom>" +
                "<a:solidFill><a:prstClr val=\"white\" /></a:solidFill><a:ln w=\"1\"><a:solidFill><a:prstClr val=\"green\" /></a:solidFill></a:ln></xdr:spPr>" +
                "<xdr:txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\" /><a:lstStyle /><a:p><a:r><a:rPr lang=\"en-US\" sz=\"1100\" />" +
                "<a:t>This shape represents a slicer. Slicers are supported in Excel 2010 or later.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>";
            node.AppendChild(doc.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings));
            SetSize(192, 240);
            WriteCacheItems();
        }

        /// <summary>A slicer of a loaded workbook.</summary>
        ExcelPivotTableSlicer(ExcelDrawings drawings, XmlNode node, ExcelTableSlicer.SlicerPart slicerPart, XmlElement slicer, ExcelTableSlicer.SlicerCache cache)
            : base(drawings, node, ExcelTableSlicer.NamePath)
        {
            _slicerPart = slicerPart;
            _slicer = slicer;
            _cache = cache;
        }

        /// <summary>
        /// The slicer of a pivot table in a drawing of a loaded workbook, or null when the drawing is not one (its slicer cache
        /// names no pivot table).
        /// </summary>
        internal static ExcelPivotTableSlicer TryLoad(ExcelDrawings drawings, XmlNode node)
        {
            WithNamespaces(drawings, node);
            XmlElement graphic = node.SelectSingleNode(ExcelTableSlicer.GraphicPath, drawings.NameSpaceManager) as XmlElement;
            if (graphic == null) return null;
            string name = graphic.GetAttribute("name");
            ExcelWorksheet ws = drawings.Worksheet;
            foreach (ExcelTableSlicer.SlicerPart part in ExcelTableSlicer.SlicerPart.All(ws))
            {
                foreach (XmlElement slicer in part.Xml.DocumentElement.ChildNodes.OfElements("slicer", ExcelTableSlicer.NsX14))
                {
                    if (slicer.GetAttribute("name") != name) continue;
                    ExcelTableSlicer.SlicerCache cache = ExcelTableSlicer.SlicerCache.Find(ws.Workbook, slicer.GetAttribute("cache"));
                    if (cache == null || PivotTableElement(cache) == null) return null;
                    return new ExcelPivotTableSlicer(drawings, node, part, slicer, cache);
                }
            }
            return null;
        }

        static XmlNode WithNamespaces(ExcelDrawings drawings, XmlNode node)
        {
            if (!drawings.NameSpaceManager.HasNamespace("mc")) drawings.NameSpaceManager.AddNamespace("mc", ExcelPackage.schemaMarkupCompatibility);
            if (!drawings.NameSpaceManager.HasNamespace("sle")) drawings.NameSpaceManager.AddNamespace("sle", ExcelTableSlicer.NsSle);
            return node;
        }

        static XmlElement PivotTableElement(ExcelTableSlicer.SlicerCache cache)
        {
            return cache.Xml.DocumentElement.SelectSingleNode("x14:pivotTables/x14:pivotTable", ExcelTableSlicer.NewNs(cache.Xml.NameTable)) as XmlElement;
        }

        /// <summary>
        /// The id of the pivot cache that a slicer cache names (x14:pivotCacheDefinition in the extLst of the cache): the one
        /// it has, or a new one that no pivot cache of the workbook has. Written to the copy of the cache of every pivot
        /// table that has it.
        /// </summary>
        static string PivotCacheId(ExcelPivotTable pivot)
        {
            XmlNamespaceManager ns = ExcelTableSlicer.NewNs(pivot.CacheDefinition.CacheDefinitionXml.NameTable);
            string path = "d:extLst/d:ext[@uri='" + PivotCacheExtUri + "']/x14:pivotCacheDefinition/@pivotCacheId";
            XmlNode has = pivot.CacheDefinition.TopNode.SelectSingleNode(path, ns);
            if (has != null && has.Value.Length > 0) return has.Value;

            long max = 0;
            foreach (ExcelWorksheet sheet in pivot.WorkSheet.Workbook.Worksheets)
            {
                if (sheet is ExcelChartsheet) continue;
                foreach (ExcelPivotTable t in sheet.PivotTables)
                {
                    XmlNode other = t.CacheDefinition.TopNode.SelectSingleNode(path, ExcelTableSlicer.NewNs(t.CacheDefinition.CacheDefinitionXml.NameTable));
                    long n;
                    if (other != null && long.TryParse(other.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out n) && n > max) max = n;
                }
            }
            string id = (max + 1).ToString(CultureInfo.InvariantCulture);
            foreach (ExcelPivotTable t in ExcelPivotTableFieldCollection.TablesOfCache(pivot))
            {
                XmlDocument doc = t.CacheDefinition.CacheDefinitionXml;
                XmlElement root = doc.DocumentElement;
                XmlElement extLst = root.SelectSingleNode("d:extLst", t.NameSpaceManager) as XmlElement;
                if (extLst == null)
                {
                    extLst = doc.CreateElement("extLst", ExcelPackage.schemaMain);
                    root.AppendChild(extLst); // the last child of pivotCacheDefinition
                }
                XmlElement ext = doc.CreateElement("ext", ExcelPackage.schemaMain);
                ext.SetAttribute("uri", PivotCacheExtUri);
                ext.SetAttribute("xmlns:x14", ExcelTableSlicer.NsX14);
                XmlElement def = doc.CreateElement("x14", "pivotCacheDefinition", ExcelTableSlicer.NsX14);
                def.SetAttribute("pivotCacheId", id);
                ext.AppendChild(def);
                extLst.InsertBefore(ext, extLst.FirstChild);
            }
            return id;
        }

        /// <summary>
        /// Excel only keeps a slicer of a pivot table whose pivot table and cache say they are of Excel 2010 or later; for one
        /// that says Excel 2007, as this library writes them, it drops the slicer when it opens the workbook, without a word
        /// (tried: the pivot table needs updatedVersion 4 and minRefreshableVersion, the cache createdVersion and
        /// refreshedVersion 4 and minRefreshableVersion; either without the other is not enough). So the pivot tables of the
        /// cache and every copy of the cache are raised to that, where they say less.
        /// </summary>
        static void MarkAsExcel2010(ExcelPivotTable pivot)
        {
            foreach (ExcelPivotTable t in ExcelPivotTableFieldCollection.TablesOfCache(pivot))
            {
                XmlElement table = (XmlElement)t.TopNode;
                AtLeast(table, "updatedVersion", 4);
                AtLeast(table, "createdVersion", 4);
                if (!table.HasAttribute("minRefreshableVersion")) table.SetAttribute("minRefreshableVersion", "3");
                XmlElement cache = (XmlElement)t.CacheDefinition.TopNode;
                AtLeast(cache, "createdVersion", 4);
                AtLeast(cache, "refreshedVersion", 4);
                if (!cache.HasAttribute("minRefreshableVersion")) cache.SetAttribute("minRefreshableVersion", "3");
            }
        }

        static void AtLeast(XmlElement element, string attribute, int version)
        {
            int has;
            if (!int.TryParse(element.GetAttribute(attribute), NumberStyles.Integer, CultureInfo.InvariantCulture, out has) || has < version)
            {
                element.SetAttribute(attribute, version.ToString(CultureInfo.InvariantCulture));
            }
        }

        /// <summary>The name of the slicer in the workbook (the slicer element); the drawing has it as its name too.</summary>
        public string SlicerName { get { return _slicer.GetAttribute("name"); } }

        /// <summary>The name of the slicer cache: the defined name Excel uses for the slicer.</summary>
        public string CacheName { get { return _cache.Name; } }

        /// <summary>The pivot table of the slicer, or null if the workbook has no pivot table with the sheet and name of the cache.</summary>
        public ExcelPivotTable PivotTable
        {
            get
            {
                if (_field != null) return _field._table;
                XmlElement e = PivotTableElement(_cache);
                if (e == null) return null;
                foreach (ExcelWorksheet sheet in _drawings.Worksheet.Workbook.Worksheets)
                {
                    if (sheet is ExcelChartsheet) continue;
                    if (sheet.SheetID.ToString(CultureInfo.InvariantCulture) != e.GetAttribute("tabId")) continue;
                    foreach (ExcelPivotTable t in sheet.PivotTables)
                    {
                        if (t.Name == e.GetAttribute("name")) return t;
                    }
                }
                return null;
            }
        }

        /// <summary>The field of the pivot table that the slicer filters, or null if the pivot table has no field with the name of the cache.</summary>
        public ExcelPivotTableField Field
        {
            get
            {
                if (_field == null)
                {
                    ExcelPivotTable pivot = PivotTable;
                    if (pivot != null) _field = pivot.Fields[_cache.Xml.DocumentElement.GetAttribute("sourceName")];
                }
                return _field;
            }
        }

        /// <summary>The text in the header of the slicer. The name of the field when the slicer is made.</summary>
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

        /// <summary>The name of the slicer style (SlicerStyleLight1 and so on), or null for the default of the workbook.</summary>
        public string Style
        {
            get
            {
                string style = _slicer.GetAttribute("style");
                return style.Length == 0 ? null : style;
            }
            set { SetSlicerAttribute("style", string.IsNullOrEmpty(value) ? null : value); }
        }

        void SetSlicerAttribute(string name, string value)
        {
            if (value == null) _slicer.RemoveAttribute(name);
            else _slicer.SetAttribute(name, value);
            _slicerPart.Write();
        }

        ExcelPivotTableField RequireField()
        {
            ExcelPivotTableField field = Field;
            if (field == null) throw new InvalidOperationException("The pivot table or the field of the slicer is not in the workbook");
            return field;
        }

        /// <summary>
        /// The values of the items that are selected: the items of the field that the pivot table shows. Empty when every item
        /// is (no filter).
        /// </summary>
        public object[] SelectedItems
        {
            get { return RequireField().ShownItemsWhenFiltered(); }
        }

        /// <summary>
        /// Selects these items and no other: the pivot table shows only them
        /// (<see cref="ExcelPivotTableField.ShowItems(object[])"/>).
        /// </summary>
        public void Select(params object[] items)
        {
            RequireField().ShowItems(items);
            WriteCacheItems();
        }

        /// <summary>Selects every item: the items of the field are all shown.</summary>
        public void ClearSelection()
        {
            RequireField().ShowAllItems();
            WriteCacheItems();
        }

        /// <summary>
        /// Writes the pivot table and the items of the slicer cache from the pivot field, as they are now: an item for each
        /// item of the field, selected (s="1") when the field shows it. Called when the slicer is made or its selection set,
        /// and before a save, because the items can have been set on the field.
        /// </summary>
        internal void WriteCacheItems()
        {
            ExcelPivotTableField field = Field;
            if (field == null) return;
            XmlNamespaceManager ns = ExcelTableSlicer.NewNs(_cache.Xml.NameTable);
            XmlElement pivotTable = PivotTableElement(_cache);
            var before = _cache.Xml.DocumentElement.OuterXml;
            if (pivotTable != null && ReferenceEquals(PivotTable, field._table))
            {
                // the pivot table can have been renamed
                if (pivotTable.GetAttribute("name") != field._table.Name) pivotTable.SetAttribute("name", field._table.Name);
            }
            XmlElement tabular = _cache.Xml.DocumentElement.SelectSingleNode("x14:data/x14:tabular", ns) as XmlElement;
            if (tabular == null) return; // an OLAP slicer: not written here
            List<KeyValuePair<int, bool>> items = field.ItemsForSlicer();
            XmlElement old = tabular.SelectSingleNode("x14:items", ns) as XmlElement;
            XmlElement list = _cache.Xml.CreateElement("items", ExcelTableSlicer.NsX14);
            list.SetAttribute("count", items.Count.ToString(CultureInfo.InvariantCulture));
            foreach (KeyValuePair<int, bool> i in items)
            {
                XmlElement e = _cache.Xml.CreateElement("i", ExcelTableSlicer.NsX14);
                e.SetAttribute("x", i.Key.ToString(CultureInfo.InvariantCulture));
                if (!i.Value) e.SetAttribute("s", "1");
                list.AppendChild(e);
            }
            if (old != null)
            {
                // the same items and selection: left as the file has them (the order and the "no data" marks of Excel)
                if (SameItems(old, items))
                {
                    if (_cache.Xml.DocumentElement.OuterXml != before) ExcelTableSlicer.WriteXml(_cache.Part, _cache.Xml); // the name of the pivot table
                    return;
                }
                tabular.ReplaceChild(list, old);
            }
            else
            {
                tabular.InsertBefore(list, tabular.FirstChild);
            }
            if (_cache.Xml.DocumentElement.OuterXml != before) ExcelTableSlicer.WriteXml(_cache.Part, _cache.Xml);
        }

        static bool SameItems(XmlElement old, List<KeyValuePair<int, bool>> items)
        {
            var selected = new Dictionary<int, bool>();
            foreach (XmlNode n in old.ChildNodes)
            {
                XmlElement e = n as XmlElement;
                if (e == null || e.LocalName != "i") continue;
                int x;
                string xs = e.GetAttribute("x");
                if (xs.Length == 0) x = 0;
                else if (!int.TryParse(xs, NumberStyles.Integer, CultureInfo.InvariantCulture, out x)) return false;
                string s = e.GetAttribute("s");
                selected[x] = s == "1" || s == "true";
            }
            if (selected.Count != items.Count) return false;
            foreach (KeyValuePair<int, bool> i in items)
            {
                bool s;
                if (!selected.TryGetValue(i.Key, out s) || s == i.Value) return false; // i.Value is "hidden"
            }
            return true;
        }

        /// <summary>
        /// Removed from the worksheet: the drawing, the slicer element (with the slicer part when it was the last of its kind
        /// on the sheet), and the slicer cache with its defined name when no other slicer uses it.
        /// </summary>
        internal override void DeleteMe()
        {
            ExcelWorksheet ws = _drawings.Worksheet;
            _slicerPart.Xml.DocumentElement.RemoveChild(_slicer);
            if (_slicerPart.Xml.DocumentElement.ChildNodes.OfElements("slicer", ExcelTableSlicer.NsX14).Count == 0) _slicerPart.Delete(ws);
            else _slicerPart.Write();
            bool used = false;
            foreach (ExcelWorksheet sheet in ws.Workbook.Worksheets)
            {
                if (sheet is ExcelChartsheet) continue;
                foreach (ExcelTableSlicer.SlicerPart part in ExcelTableSlicer.SlicerPart.All(sheet))
                {
                    foreach (XmlElement s in part.Xml.DocumentElement.ChildNodes.OfElements("slicer", ExcelTableSlicer.NsX14))
                    {
                        if (s.GetAttribute("cache") == _cache.Name) used = true;
                    }
                }
            }
            if (!used) _cache.Delete(ws.Workbook);
            base.DeleteMe();
        }
    }
}
