/*******************************************************************************
 * orabooster-epplus: this file is new in the fork (item B2 of issue #8, issue #56). It is not part of EPPlus 4.5 or EPPlusFree
 * and has no code from EPPlus 5 or later. Written by OraScope Co.,Ltd. 2026 from ECMA-376 Part 1 (18.10.1.1 autoSortScope,
 * 18.3.1.68 pivotArea, 18.3.1.72 reference) and from files Excel made, under the license of the library (LGPL-3.0).
 *
 * The items of a pivot field sorted by the values of a data field:
 *
 *   var sales = pivot.DataFields.Add(pivot.Fields["Sales"]);
 *   pivot.Fields["Region"].SetSort(eSortType.Descending, sales);     // the regions with the largest sales first
 *
 * In the file: the sortType of the pivot field and an autoSortScope with a pivotArea that names the data field by its
 * position in dataFields. The position is written on save, from the data field that was given.
 *******************************************************************************/
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    public partial class ExcelPivotTableField
    {
        const string DataFieldReference = "4294967294"; // the "field" of a reference that stands for the data fields (-2)

        ExcelPivotTableDataField _sortDataField;
        bool _sortDataFieldRead;

        /// <summary>
        /// The data field whose values the items of this field are sorted by, or null when they are sorted by their labels (or
        /// not sorted). Set with <see cref="SetSort(eSortType, ExcelPivotTableDataField)"/>.
        /// </summary>
        public ExcelPivotTableDataField SortDataField
        {
            get
            {
                if (!_sortDataFieldRead)
                {
                    _sortDataFieldRead = true;
                    XmlElement x = TopNode.SelectSingleNode("d:autoSortScope/d:pivotArea/d:references/d:reference[@field='" + DataFieldReference + "']/d:x", NameSpaceManager) as XmlElement;
                    if (x != null)
                    {
                        int index;
                        string v = x.GetAttribute("v"); // 0 when it is left out
                        if (v.Length == 0) index = 0;
                        else if (!int.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out index)) index = -1;
                        if (index >= 0 && index < _table.DataFields.Count) _sortDataField = _table.DataFields[index];
                    }
                }
                return _sortDataField;
            }
        }

        /// <summary>
        /// Sorts the items of this field by the values of a data field of the pivot table, as "More sort options" of Excel
        /// does: ascending or descending. Excel sorts when it refreshes the pivot table, which it does when the workbook is
        /// opened. With null as the data field the items are sorted by their labels, as <see cref="Sort"/> does.
        /// </summary>
        /// <param name="sort">Ascending or descending; None takes the sort off</param>
        /// <param name="dataField">A data field of this pivot table, or null</param>
        public void SetSort(eSortType sort, ExcelPivotTableDataField dataField)
        {
            if (dataField != null)
            {
                if (sort == eSortType.None) throw new ArgumentException("A sort by a data field is ascending or descending", "sort");
                bool found = false;
                foreach (ExcelPivotTableDataField d in _table.DataFields)
                {
                    if (ReferenceEquals(d, dataField)) found = true;
                }
                if (!found) throw new ArgumentException("The data field is not a data field of this pivot table", "dataField");
            }
            Sort = sort; // takes the scope away
            _sortDataField = dataField;
            _sortDataFieldRead = true;
            WriteAutoSort();
            RefreshOnLoad(_table);
        }

        /// <summary>
        /// Excel is to refresh the pivot table when it opens the workbook: a pivot table of a loaded workbook has its rows as
        /// they were saved, and shows a change (a sort, a filter) only after a refresh. A pivot table made here has this set
        /// from the start. Set in the copy of the cache of every pivot table that has the cache.
        /// </summary>
        internal static void RefreshOnLoad(ExcelPivotTable pivotTable)
        {
            foreach (ExcelPivotTable t in ExcelPivotTableFieldCollection.TablesOfCache(pivotTable))
            {
                ((XmlElement)t.CacheDefinition.TopNode).SetAttribute("refreshOnLoad", "1");
            }
        }

        /// <summary>Called when <see cref="Sort"/> is set: the items are sorted by their labels, not by a data field.</summary>
        void ClearAutoSort()
        {
            XmlNode scope = TopNode.SelectSingleNode("d:autoSortScope", NameSpaceManager);
            if (scope != null) TopNode.RemoveChild(scope);
            _sortDataField = null;
            _sortDataFieldRead = true;
        }

        /// <summary>
        /// Writes the autoSortScope from the data field that was given: when it is set, and again before a save, because the
        /// position of the data field can have changed. A data field that is no longer in the pivot table takes the scope
        /// away (the items are then sorted by their labels). A scope of a loaded workbook that was not asked for is not touched.
        /// </summary>
        internal void WriteAutoSort()
        {
            if (!_sortDataFieldRead) return;
            XmlNode old = TopNode.SelectSingleNode("d:autoSortScope", NameSpaceManager);
            int index = -1;
            if (_sortDataField != null)
            {
                for (int i = 0; i < _table.DataFields.Count; i++)
                {
                    if (ReferenceEquals(_table.DataFields[i], _sortDataField)) index = i;
                }
            }
            if (index < 0)
            {
                if (_sortDataField != null)
                {
                    // the data field was taken out of the pivot table: nothing to sort by
                    if (old != null) TopNode.RemoveChild(old);
                    _sortDataField = null;
                }
                return; // no data field: a scope of the file (of another kind, or one that could not be read) stays as it is
            }
            if (old != null)
            {
                XmlElement x = old.SelectSingleNode("d:pivotArea/d:references/d:reference[@field='" + DataFieldReference + "']/d:x", NameSpaceManager) as XmlElement;
                if (x != null)
                {
                    string v = x.GetAttribute("v");
                    int was = v.Length == 0 ? 0 : -1;
                    if (was < 0) int.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out was);
                    if (was == index) return; // as it is
                }
                TopNode.RemoveChild(old);
            }
            XmlDocument doc = TopNode.OwnerDocument;
            XmlElement scope = doc.CreateElement("autoSortScope", ExcelPackage.schemaMain);
            scope.InnerXml = "<pivotArea xmlns=\"" + ExcelPackage.schemaMain + "\" dataOnly=\"0\" outline=\"0\" fieldPosition=\"0\"><references count=\"1\">" +
                "<reference field=\"" + DataFieldReference + "\" count=\"1\" selected=\"0\"><x v=\"" + index.ToString(CultureInfo.InvariantCulture) + "\" /></reference></references></pivotArea>";
            // CT_PivotField: items, autoSortScope, extLst
            XmlNode items = TopNode.SelectSingleNode("d:items", NameSpaceManager);
            if (items != null) TopNode.InsertAfter(scope, items);
            else TopNode.InsertBefore(scope, TopNode.FirstChild);
        }
    }
}
