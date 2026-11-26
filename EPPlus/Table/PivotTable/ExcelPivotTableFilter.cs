/*******************************************************************************
 * orabooster-epplus: this file is new in the fork (item B2 of issue #8, issue #60). It is not part of EPPlus 4.5 or EPPlusFree
 * and has no code from EPPlus 5 or later. Written by OraScope Co.,Ltd. 2026 from ECMA-376 Part 1 (18.10.1.33 filter,
 * 18.10.1.34 filters, 18.18.59 ST_PivotFilterType, and the autoFilter of 18.3.2) and from a workbook in which Excel made each
 * kind of filter, under the license of the library (LGPL-3.0).
 *
 * Label filters and value filters of a pivot field:
 *
 *   pivot.Fields["Product"].SetLabelFilter(ePivotLabelFilter.BeginsWith, "P1");
 *   pivot.Fields["Region"].SetValueFilter(ePivotValueFilter.GreaterThan, salesDataField, 100);
 *   pivot.Fields["Region"].SetValueFilter(ePivotValueFilter.TopCount, salesDataField, 3);
 *
 * In the file: a filter element in the filters of the pivot table, with the kind, the values, and an autoFilter that says
 * the same in the form of a worksheet filter. Excel applies it when it refreshes the pivot table (on load).
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>The kinds of label filter of a pivot field: the items are compared by their text.</summary>
    public enum ePivotLabelFilter
    {
        /// <summary>The text of the item is the value</summary>
        Equals,
        /// <summary>The text of the item is not the value</summary>
        NotEquals,
        /// <summary>Begins with the value</summary>
        BeginsWith,
        /// <summary>Does not begin with the value</summary>
        NotBeginsWith,
        /// <summary>Ends with the value</summary>
        EndsWith,
        /// <summary>Does not end with the value</summary>
        NotEndsWith,
        /// <summary>Contains the value</summary>
        Contains,
        /// <summary>Does not contain the value</summary>
        NotContains,
        /// <summary>Is after the value</summary>
        GreaterThan,
        /// <summary>Is the value or after it</summary>
        GreaterThanOrEqual,
        /// <summary>Is before the value</summary>
        LessThan,
        /// <summary>Is the value or before it</summary>
        LessThanOrEqual,
        /// <summary>Is between the two values, or one of them</summary>
        Between,
        /// <summary>Is before the first value or after the second</summary>
        NotBetween
    }

    /// <summary>The kinds of value filter of a pivot field: the items are compared by their value in a data field.</summary>
    public enum ePivotValueFilter
    {
        /// <summary>The value is the number</summary>
        Equals,
        /// <summary>The value is not the number</summary>
        NotEquals,
        /// <summary>Greater than the number</summary>
        GreaterThan,
        /// <summary>The number or greater</summary>
        GreaterThanOrEqual,
        /// <summary>Less than the number</summary>
        LessThan,
        /// <summary>The number or less</summary>
        LessThanOrEqual,
        /// <summary>Between the two numbers, or one of them</summary>
        Between,
        /// <summary>Less than the first number or greater than the second</summary>
        NotBetween,
        /// <summary>The items with the largest values: as many as the number says</summary>
        TopCount,
        /// <summary>The items with the smallest values: as many as the number says</summary>
        BottomCount,
        /// <summary>The items with the largest values that make up the percentage of the total</summary>
        TopPercent,
        /// <summary>The items with the smallest values that make up the percentage of the total</summary>
        BottomPercent,
        /// <summary>The items with the largest values whose sum reaches the number</summary>
        TopSum,
        /// <summary>The items with the smallest values whose sum reaches the number</summary>
        BottomSum
    }

    /// <summary>
    /// A label filter or a value filter of a pivot field, as it is in the file. Read from
    /// <see cref="ExcelPivotTableField.LabelFilter"/> and <see cref="ExcelPivotTableField.ValueFilter"/>.
    /// </summary>
    public sealed class ExcelPivotTableFilter
    {
        internal ExcelPivotTableFilter() { }

        /// <summary>True for a value filter, false for a label filter.</summary>
        public bool IsValueFilter { get; internal set; }
        /// <summary>The kind of a label filter; null for a value filter, or for a kind this library has no name for (a date filter).</summary>
        public ePivotLabelFilter? LabelType { get; internal set; }
        /// <summary>The kind of a value filter; null for a label filter.</summary>
        public ePivotValueFilter? ValueType { get; internal set; }
        /// <summary>The kind as the file has it (ST_PivotFilterType), such as captionBeginsWith or valueGreaterThan.</summary>
        public string TypeName { get; internal set; }
        /// <summary>The first text of a label filter.</summary>
        public string Text1 { get; internal set; }
        /// <summary>The second text of a label filter (between, not between), or null.</summary>
        public string Text2 { get; internal set; }
        /// <summary>The first number of a value filter.</summary>
        public double? Value1 { get; internal set; }
        /// <summary>The second number of a value filter (between, not between), or null.</summary>
        public double? Value2 { get; internal set; }
        /// <summary>The data field of a value filter, or null.</summary>
        public ExcelPivotTableDataField DataField { get; internal set; }
    }

    public partial class ExcelPivotTableField
    {
        static readonly string[] _labelTypes = { "captionEqual", "captionNotEqual", "captionBeginsWith", "captionNotBeginsWith", "captionEndsWith", "captionNotEndsWith",
            "captionContains", "captionNotContains", "captionGreaterThan", "captionGreaterThanOrEqual", "captionLessThan", "captionLessThanOrEqual", "captionBetween", "captionNotBetween" };
        static readonly string[] _valueTypes = { "valueEqual", "valueNotEqual", "valueGreaterThan", "valueGreaterThanOrEqual", "valueLessThan", "valueLessThanOrEqual",
            "valueBetween", "valueNotBetween", "count", "count", "percent", "percent", "sum", "sum" };

        ExcelPivotTableDataField _valueFilterDataField; // the data field of a value filter set here: its position is written on save

        static bool IsValueType(string type)
        {
            return type.StartsWith("value", StringComparison.Ordinal) || type == "count" || type == "percent" || type == "sum";
        }

        /// <summary>The filter element of this field in the filters of the pivot table: the label filter or the value filter.</summary>
        XmlElement FilterElement(bool valueFilter)
        {
            foreach (XmlElement f in _table.TopNode.SelectNodes("d:filters/d:filter[@fld='" + Index.ToString(CultureInfo.InvariantCulture) + "']", NameSpaceManager))
            {
                if (IsValueType(f.GetAttribute("type")) == valueFilter) return f;
            }
            return null;
        }

        ExcelPivotTableFilter ReadFilter(XmlElement f)
        {
            if (f == null) return null;
            var filter = new ExcelPivotTableFilter();
            string type = f.GetAttribute("type");
            filter.TypeName = type;
            filter.IsValueFilter = IsValueType(type);
            if (filter.IsValueFilter)
            {
                XmlElement top10 = f.SelectSingleNode("d:autoFilter/d:filterColumn/d:top10", NameSpaceManager) as XmlElement;
                if (top10 != null)
                {
                    bool bottom = top10.GetAttribute("top") == "0";
                    filter.ValueType = type == "count" ? (bottom ? ePivotValueFilter.BottomCount : ePivotValueFilter.TopCount)
                        : type == "percent" ? (bottom ? ePivotValueFilter.BottomPercent : ePivotValueFilter.TopPercent)
                        : (bottom ? ePivotValueFilter.BottomSum : ePivotValueFilter.TopSum);
                    filter.Value1 = ToNumber(top10.GetAttribute("val"));
                }
                else
                {
                    int i = Array.IndexOf(_valueTypes, type);
                    if (i >= 0) filter.ValueType = (ePivotValueFilter)i;
                    var custom = f.SelectNodes("d:autoFilter/d:filterColumn/d:customFilters/d:customFilter", NameSpaceManager);
                    if (custom.Count > 0) filter.Value1 = ToNumber(((XmlElement)custom[0]).GetAttribute("val"));
                    if (custom.Count > 1) filter.Value2 = ToNumber(((XmlElement)custom[1]).GetAttribute("val"));
                }
                int measure;
                if (int.TryParse(f.GetAttribute("iMeasureFld"), NumberStyles.Integer, CultureInfo.InvariantCulture, out measure) && measure >= 0 && measure < _table.DataFields.Count)
                {
                    filter.DataField = _table.DataFields[measure];
                }
                if (_valueFilterDataField != null) filter.DataField = _valueFilterDataField;
            }
            else
            {
                int i = Array.IndexOf(_labelTypes, type);
                if (i >= 0) filter.LabelType = (ePivotLabelFilter)i;
                filter.Text1 = f.HasAttribute("stringValue1") ? f.GetAttribute("stringValue1") : null;
                filter.Text2 = f.HasAttribute("stringValue2") ? f.GetAttribute("stringValue2") : null;
            }
            return filter;
        }

        static double? ToNumber(string text)
        {
            double d;
            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out d) ? d : (double?)null;
        }

        /// <summary>The label filter of this field, or null. Set with <see cref="SetLabelFilter(ePivotLabelFilter, string, string)"/>.</summary>
        public ExcelPivotTableFilter LabelFilter
        {
            get { return ReadFilter(FilterElement(false)); }
        }

        /// <summary>The value filter of this field, or null. Set with <see cref="SetValueFilter(ePivotValueFilter, ExcelPivotTableDataField, double, double?)"/>.</summary>
        public ExcelPivotTableFilter ValueFilter
        {
            get { return ReadFilter(FilterElement(true)); }
        }

        /// <summary>
        /// Sets the label filter of this field: the pivot table shows the items whose text meets the condition. The text is
        /// compared without case; it is taken as it is (a star or a question mark is that character, not a wildcard). A field
        /// has one label filter; this replaces it. With a value filter on the same field too, the pivot table is set to allow
        /// several filters for a field. Excel applies the filter when it opens the workbook.
        /// </summary>
        /// <param name="type">The kind of filter</param>
        /// <param name="text1">The text</param>
        /// <param name="text2">The second text, for between and not between</param>
        public void SetLabelFilter(ePivotLabelFilter type, string text1, string text2 = null)
        {
            if (text1 == null) throw new ArgumentNullException("text1");
            bool two = type == ePivotLabelFilter.Between || type == ePivotLabelFilter.NotBetween;
            if (two && text2 == null) throw new ArgumentNullException("text2", "Between and not between take two texts");
            if (!two && text2 != null) throw new ArgumentException("Only between and not between take a second text", "text2");
            if (IsCalculated) throw new InvalidOperationException("A calculated field has no items to filter");

            XmlElement filter = NewFilterElement(false, _labelTypes[(int)type]);
            filter.SetAttribute("stringValue1", text1);
            if (two) filter.SetAttribute("stringValue2", text2);
            XmlElement column = AddAutoFilter(filter);
            string a = EscapeWildcards(text1), b = two ? EscapeWildcards(text2) : null;
            switch (type)
            {
                case ePivotLabelFilter.Equals:
                    XmlElement filters = Append(column, "filters");
                    Append(filters, "filter").SetAttribute("val", text1);
                    break;
                case ePivotLabelFilter.NotEquals: Custom(column, false, "notEqual", a); break;
                case ePivotLabelFilter.BeginsWith: Custom(column, false, null, a + "*"); break;
                case ePivotLabelFilter.NotBeginsWith: Custom(column, false, "notEqual", a + "*"); break;
                case ePivotLabelFilter.EndsWith: Custom(column, false, null, "*" + a); break;
                case ePivotLabelFilter.NotEndsWith: Custom(column, false, "notEqual", "*" + a); break;
                case ePivotLabelFilter.Contains: Custom(column, false, null, "*" + a + "*"); break;
                case ePivotLabelFilter.NotContains: Custom(column, false, "notEqual", "*" + a + "*"); break;
                case ePivotLabelFilter.GreaterThan: Custom(column, false, "greaterThan", a); break;
                case ePivotLabelFilter.GreaterThanOrEqual: Custom(column, false, "greaterThanOrEqual", a); break;
                case ePivotLabelFilter.LessThan: Custom(column, false, "lessThan", a); break;
                case ePivotLabelFilter.LessThanOrEqual: Custom(column, false, "lessThanOrEqual", a); break;
                case ePivotLabelFilter.Between: Custom(column, true, "greaterThanOrEqual", a, "lessThanOrEqual", b); break;
                case ePivotLabelFilter.NotBetween: Custom(column, false, "lessThan", a, "greaterThan", b); break;
            }
            FiltersChanged();
        }

        /// <summary>
        /// Sets the value filter of this field: the pivot table shows the items whose value in a data field meets the
        /// condition, or the top or bottom items by that value. A field has one value filter; this replaces it. Excel applies
        /// the filter when it opens the workbook.
        /// </summary>
        /// <param name="type">The kind of filter</param>
        /// <param name="dataField">A data field of this pivot table</param>
        /// <param name="value1">The number: the value to compare with, the number of items, the percentage or the sum</param>
        /// <param name="value2">The second number, for between and not between</param>
        public void SetValueFilter(ePivotValueFilter type, ExcelPivotTableDataField dataField, double value1, double? value2 = null)
        {
            if (dataField == null) throw new ArgumentNullException("dataField");
            if (DataFieldIndex(dataField) < 0) throw new ArgumentException("The data field is not a data field of this pivot table", "dataField");
            bool two = type == ePivotValueFilter.Between || type == ePivotValueFilter.NotBetween;
            if (two && value2 == null) throw new ArgumentNullException("value2", "Between and not between take two numbers");
            if (!two && value2 != null) throw new ArgumentException("Only between and not between take a second number", "value2");
            if (double.IsNaN(value1) || double.IsInfinity(value1) || (two && (double.IsNaN(value2.Value) || double.IsInfinity(value2.Value)))) throw new ArgumentException("The number of a value filter is a finite number");
            bool top = type >= ePivotValueFilter.TopCount;
            if (top && value1 <= 0) throw new ArgumentOutOfRangeException("value1", "The number of a top or bottom filter is more than 0");
            if ((type == ePivotValueFilter.TopCount || type == ePivotValueFilter.BottomCount) && value1 != Math.Floor(value1)) throw new ArgumentOutOfRangeException("value1", "The number of items is a whole number");
            if ((type == ePivotValueFilter.TopPercent || type == ePivotValueFilter.BottomPercent) && value1 > 100) throw new ArgumentOutOfRangeException("value1", "A percentage is at most 100");
            if (IsCalculated) throw new InvalidOperationException("A calculated field has no items to filter");

            XmlElement filter = NewFilterElement(true, _valueTypes[(int)type]);
            _valueFilterDataField = dataField;
            filter.SetAttribute("iMeasureFld", DataFieldIndex(dataField).ToString(CultureInfo.InvariantCulture));
            XmlElement column = AddAutoFilter(filter);
            string a = Number(value1), b = two ? Number(value2.Value) : null;
            switch (type)
            {
                case ePivotValueFilter.Equals: Custom(column, false, null, a); break;
                case ePivotValueFilter.NotEquals: Custom(column, false, "notEqual", a); break;
                case ePivotValueFilter.GreaterThan: Custom(column, false, "greaterThan", a); break;
                case ePivotValueFilter.GreaterThanOrEqual: Custom(column, false, "greaterThanOrEqual", a); break;
                case ePivotValueFilter.LessThan: Custom(column, false, "lessThan", a); break;
                case ePivotValueFilter.LessThanOrEqual: Custom(column, false, "lessThanOrEqual", a); break;
                case ePivotValueFilter.Between: Custom(column, true, "greaterThanOrEqual", a, "lessThanOrEqual", b); break;
                case ePivotValueFilter.NotBetween: Custom(column, false, "lessThan", a, "greaterThan", b); break;
                default:
                    XmlElement top10 = Append(column, "top10");
                    if (type == ePivotValueFilter.BottomCount || type == ePivotValueFilter.BottomPercent || type == ePivotValueFilter.BottomSum) top10.SetAttribute("top", "0");
                    if (type == ePivotValueFilter.TopPercent || type == ePivotValueFilter.BottomPercent) top10.SetAttribute("percent", "1");
                    top10.SetAttribute("val", a);
                    top10.SetAttribute("filterVal", a);
                    break;
            }
            SetXmlNodeString("@measureFilter", "1");
            FiltersChanged();
        }

        /// <summary>Takes the label filter of this field away.</summary>
        public void ClearLabelFilter()
        {
            if (RemoveFilterElement(false)) FiltersChanged();
        }

        /// <summary>Takes the value filter of this field away.</summary>
        public void ClearValueFilter()
        {
            if (RemoveFilterElement(true)) FiltersChanged();
        }

        int DataFieldIndex(ExcelPivotTableDataField dataField)
        {
            for (int i = 0; i < _table.DataFields.Count; i++)
            {
                if (ReferenceEquals(_table.DataFields[i], dataField)) return i;
            }
            return -1;
        }

        static string Number(double value)
        {
            return value.ToString("R", CultureInfo.InvariantCulture);
        }

        // In a customFilter a star and a question mark are wildcards and a tilde makes the next character plain
        static string EscapeWildcards(string text)
        {
            return text.Replace("~", "~~").Replace("*", "~*").Replace("?", "~?");
        }

        static XmlElement Append(XmlElement parent, string name)
        {
            XmlElement e = parent.OwnerDocument.CreateElement(name, ExcelPackage.schemaMain);
            parent.AppendChild(e);
            return e;
        }

        static void Custom(XmlElement column, bool and, string operator1, string value1, string operator2 = null, string value2 = null)
        {
            XmlElement custom = Append(column, "customFilters");
            if (and) custom.SetAttribute("and", "1");
            XmlElement first = Append(custom, "customFilter");
            if (operator1 != null) first.SetAttribute("operator", operator1);
            first.SetAttribute("val", value1);
            if (value2 != null)
            {
                XmlElement second = Append(custom, "customFilter");
                if (operator2 != null) second.SetAttribute("operator", operator2);
                second.SetAttribute("val", value2);
            }
        }

        static XmlElement AddAutoFilter(XmlElement filter)
        {
            XmlElement autoFilter = Append(filter, "autoFilter");
            autoFilter.SetAttribute("ref", "A1");
            XmlElement column = Append(autoFilter, "filterColumn");
            column.SetAttribute("colId", "0");
            return column;
        }

        bool RemoveFilterElement(bool valueFilter)
        {
            XmlElement old = FilterElement(valueFilter);
            if (old == null) return false;
            old.ParentNode.RemoveChild(old);
            if (valueFilter)
            {
                _valueFilterDataField = null;
                DeleteNode("@measureFilter");
            }
            return true;
        }

        /// <summary>A new filter element for this field, in place of the one of that sort it had.</summary>
        XmlElement NewFilterElement(bool valueFilter, string type)
        {
            RemoveFilterElement(valueFilter);
            XmlElement filters = (XmlElement)_table.CreateNode("d:filters");
            int id = 1;
            foreach (XmlElement f in filters.SelectNodes("d:filter", NameSpaceManager))
            {
                int n;
                if (int.TryParse(f.GetAttribute("id"), NumberStyles.Integer, CultureInfo.InvariantCulture, out n) && n >= id) id = n + 1;
            }
            XmlElement filter = Append(filters, "filter");
            filter.SetAttribute("fld", Index.ToString(CultureInfo.InvariantCulture));
            filter.SetAttribute("type", type);
            filter.SetAttribute("evalOrder", "-1");
            filter.SetAttribute("id", id.ToString(CultureInfo.InvariantCulture));
            return filter;
        }

        /// <summary>After a filter was set or taken away: the count, the empty filters element, several filters for a field, a refresh.</summary>
        void FiltersChanged()
        {
            XmlElement filters = _table.TopNode.SelectSingleNode("d:filters", NameSpaceManager) as XmlElement;
            if (filters != null)
            {
                int count = filters.SelectNodes("d:filter", NameSpaceManager).Count;
                if (count == 0) filters.ParentNode.RemoveChild(filters);
                else filters.SetAttribute("count", count.ToString(CultureInfo.InvariantCulture));
            }
            // a label filter and a value filter on one field: Excel asks for "allow multiple filters per field"
            if (FilterElement(false) != null && FilterElement(true) != null) _table.MultipleFieldFilters = true;
            RefreshOnLoad(_table);
        }

        /// <summary>
        /// Before a save: the position of the data field of a value filter set here, as it is now. A data field that is no
        /// longer in the pivot table takes the filter away. A filter of a loaded workbook that was not set here is not touched.
        /// </summary>
        internal void WriteValueFilter()
        {
            if (_valueFilterDataField == null) return;
            XmlElement f = FilterElement(true);
            if (f == null) { _valueFilterDataField = null; return; }
            int index = DataFieldIndex(_valueFilterDataField);
            if (index < 0)
            {
                RemoveFilterElement(true);
                FiltersChanged();
                return;
            }
            string v = index.ToString(CultureInfo.InvariantCulture);
            if (f.GetAttribute("iMeasureFld") != v) f.SetAttribute("iMeasureFld", v);
        }
    }
}
