/*******************************************************************************
 * orabooster-epplus: this file is new in the fork (item B2 of issue #8, issue #62). It is not part of EPPlus 4.5 or EPPlusFree
 * and has no code from EPPlus 5 or later. Written by OraScope Co.,Ltd. 2026 from ECMA-376 Part 1 (18.10.1.90 sharedItems,
 * 18.10.1.45 item, 18.10.1.46 items, 18.10.1.62 pageField) and from workbooks in which Excel made pivot tables over every
 * kind of value, under the license of the library (LGPL-3.0).
 *
 * A pivot field filtered by its items:
 *
 *   pivot.Fields["Region"].HideItems("North");            // every region but North
 *   pivot.Fields["Product"].ShowItems("P1", "P2");        // only these
 *   pivot.Fields["Year"].SetPageItem(2026);               // a page field that shows one item
 *
 * In the file the items of a field are the sharedItems of its cache field (each value of the column once, in the order of
 * the source), the pivot field has an item for each of them, and h="1" hides one. The library writes a pivot table with no
 * items and lets Excel make them on the refresh; for a field that is filtered here it makes them itself, from the column of
 * the source, and makes them again on save so that they follow the cells.
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Table.PivotTable
{
    public partial class ExcelPivotTableField
    {
        /// <summary>One item of a field: a value of its column of the source, as the cache has it.</summary>
        sealed class SharedItem
        {
            internal char Kind;      // s text, n number, d date, b boolean, e error, m blank
            internal string Text;    // the v attribute
            internal object Value;   // string, double, DateTime, bool, the error text, or null for blank
            internal string Key { get { return Kind == 's' ? "s|" + Text.ToLowerInvariant() : Kind + "|" + Text; } }
        }

        bool _itemsMadeHere;                 // the shared items were made from the source here: made again on save
        HashSet<string> _hiddenKeys;         // the items hidden here, by key, so that they are found again after the cells changed
        string _pageItemKey;                 // the item of the page field set here
        bool _pageItemSet;

        static SharedItem ToItem(object value, bool dateFormat)
        {
            if (value == null) return new SharedItem { Kind = 'm', Text = "" };
            if (value is string)
            {
                string s = (string)value;
                return s.Length == 0 ? new SharedItem { Kind = 'm', Text = "" } : new SharedItem { Kind = 's', Text = s, Value = s };
            }
            if (value is bool) return new SharedItem { Kind = 'b', Text = (bool)value ? "1" : "0", Value = value };
            if (value is DateTime) return DateItem((DateTime)value);
            if (value is TimeSpan) return DateItem(DateTime.FromOADate(((TimeSpan)value).TotalDays));
            if (value is ExcelErrorValue) return new SharedItem { Kind = 'e', Text = value.ToString(), Value = value.ToString() };
            if (value is double || value is float || value is decimal || value is int || value is long || value is short || value is byte
                || value is uint || value is ulong || value is ushort || value is sbyte)
            {
                double d = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                if (dateFormat)
                {
                    try { return DateItem(DateTime.FromOADate(d)); } catch (ArgumentException) { }
                }
                return new SharedItem { Kind = 'n', Text = d.ToString("R", CultureInfo.InvariantCulture), Value = d };
            }
            string text = Convert.ToString(value, CultureInfo.InvariantCulture);
            return new SharedItem { Kind = 's', Text = text, Value = text };
        }

        static SharedItem DateItem(DateTime date)
        {
            return new SharedItem { Kind = 'd', Text = date.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture), Value = date };
        }

        /// <summary>The key of a value given by the caller, as the item of that value has it.</summary>
        static string KeyOf(object value)
        {
            return ToItem(value, false).Key;
        }

        /// <summary>The items the cache field has now.</summary>
        List<SharedItem> ReadSharedItems()
        {
            var list = new List<SharedItem>();
            if (_cacheFieldHelper == null) return list;
            XmlNode shared = _cacheFieldHelper.TopNode.SelectSingleNode("d:sharedItems", NameSpaceManager);
            if (shared == null) return list;
            foreach (XmlNode node in shared.ChildNodes)
            {
                XmlElement e = node as XmlElement;
                if (e == null) continue;
                string v = e.GetAttribute("v");
                switch (e.LocalName)
                {
                    case "s": list.Add(new SharedItem { Kind = 's', Text = v, Value = v }); break;
                    case "n":
                        double d;
                        double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out d);
                        list.Add(new SharedItem { Kind = 'n', Text = v, Value = d });
                        break;
                    case "d":
                        DateTime date;
                        DateTime.TryParse(v, CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
                        list.Add(new SharedItem { Kind = 'd', Text = v, Value = date });
                        break;
                    case "b": list.Add(new SharedItem { Kind = 'b', Text = v == "1" || v == "true" ? "1" : "0", Value = v == "1" || v == "true" }); break;
                    case "e": list.Add(new SharedItem { Kind = 'e', Text = v, Value = v }); break;
                    case "m": list.Add(new SharedItem { Kind = 'm', Text = "" }); break;
                }
            }
            return list;
        }

        /// <summary>The items of the column of this field in the source: each value once, in the order of the rows.</summary>
        List<SharedItem> ItemsOfSource()
        {
            ExcelRangeBase source = _table.CacheDefinition.SourceRange;
            if (source == null) throw new InvalidOperationException("The source of the pivot table is not a range of this workbook, so the items of the field " + Name + " can not be made");
            ExcelWorksheet ws = source.Worksheet;
            int col = source._fromCol + Index;
            if (col > source._toCol) throw new InvalidOperationException("The field " + Name + " is not a column of the source");
            var list = new List<SharedItem>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            var styles = ws.Workbook.Styles;
            for (int row = source._fromRow + 1; row <= source._toRow; row++)
            {
                object value = ws.GetValueInner(row, col);
                bool dateFormat = false;
                if (value != null && !(value is string) && !(value is bool) && !(value is DateTime))
                {
                    // a number in a cell with a date format is a date, as Excel takes it (found as ExcelRangeBase finds the format of a cell)
                    int styleId = ws.Cells[row, col].StyleID;
                    if (styleId >= 0 && styleId < styles.CellXfs.Count)
                    {
                        int numFmtId = styles.CellXfs[styleId].NumberFormatId;
                        for (int i = 0; i < styles.NumberFormats.Count; i++)
                        {
                            if (styles.NumberFormats[i].NumFmtId == numFmtId)
                            {
                                dateFormat = styles.NumberFormats[i].FormatTranslator.DataType == ExcelNumberFormatXml.eFormatType.DateTime;
                                break;
                            }
                        }
                    }
                }
                SharedItem item = ToItem(value, dateFormat);
                if (seen.Add(item.Key)) list.Add(item);
            }
            return list;
        }

        /// <summary>
        /// Writes the sharedItems of the cache field, with the attributes Excel gives each mix of values (seen in workbooks
        /// Excel made: texts, booleans and errors count as text; a column with both numbers and dates is not taken).
        /// </summary>
        static void WriteSharedItems(XmlElement cacheField, List<SharedItem> items, XmlNamespaceManager ns)
        {
            bool text = false, number = false, date = false, blank = false, integers = true;
            double min = double.MaxValue, max = double.MinValue;
            DateTime minDate = DateTime.MaxValue, maxDate = DateTime.MinValue;
            foreach (SharedItem i in items)
            {
                switch (i.Kind)
                {
                    case 's': case 'b': case 'e': text = true; break;
                    case 'm': blank = true; break;
                    case 'n':
                        number = true;
                        double d = (double)i.Value;
                        if (d != Math.Floor(d)) integers = false;
                        if (d < min) min = d;
                        if (d > max) max = d;
                        break;
                    case 'd':
                        date = true;
                        DateTime t = (DateTime)i.Value;
                        if (t < minDate) minDate = t;
                        if (t > maxDate) maxDate = t;
                        break;
                }
            }
            if (number && date) throw new NotSupportedException("The column of the field " + cacheField.GetAttribute("name") + " has both numbers and dates; its items can not be made here");

            XmlDocument doc = cacheField.OwnerDocument;
            XmlElement shared = doc.CreateElement("sharedItems", ExcelPackage.schemaMain);
            if (!text && !blank) shared.SetAttribute("containsSemiMixedTypes", "0");
            if (date && !text) shared.SetAttribute("containsNonDate", "0");
            if (date) shared.SetAttribute("containsDate", "1");
            if (!text) shared.SetAttribute("containsString", "0");
            if (blank) shared.SetAttribute("containsBlank", "1");
            if (text && (number || date)) shared.SetAttribute("containsMixedTypes", "1");
            if (number)
            {
                shared.SetAttribute("containsNumber", "1");
                if (integers) shared.SetAttribute("containsInteger", "1");
                shared.SetAttribute("minValue", min.ToString("R", CultureInfo.InvariantCulture));
                shared.SetAttribute("maxValue", max.ToString("R", CultureInfo.InvariantCulture));
            }
            if (date)
            {
                shared.SetAttribute("minDate", minDate.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture));
                shared.SetAttribute("maxDate", maxDate.Date.AddDays(1).ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture)); // the day after, as Excel writes it
            }
            shared.SetAttribute("count", items.Count.ToString(CultureInfo.InvariantCulture));
            foreach (SharedItem i in items)
            {
                XmlElement e = doc.CreateElement(i.Kind.ToString(), ExcelPackage.schemaMain);
                if (i.Kind != 'm') e.SetAttribute("v", i.Text);
                shared.AppendChild(e);
            }
            XmlNode old = cacheField.SelectSingleNode("d:sharedItems", ns);
            if (old != null) cacheField.ReplaceChild(shared, old);
            else cacheField.InsertBefore(shared, cacheField.FirstChild);
        }

        /// <summary>The items element of the pivot field with an item for each shared item, in the order it has (or of the cache), and the subtotal item.</summary>
        void WritePivotItems(int count, HashSet<int> hidden)
        {
            WritePivotItems(this, count, hidden);
        }

        static void WritePivotItems(ExcelPivotTableField field, int count, HashSet<int> hidden)
        {
            XmlElement top = (XmlElement)field.TopNode;
            XmlElement old = top.SelectSingleNode("d:items", field.NameSpaceManager) as XmlElement;
            // the order the field has, as long as it names the same items
            var order = new List<int>();
            if (old != null)
            {
                foreach (XmlElement item in old.SelectNodes("d:item[not(@t)]", field.NameSpaceManager))
                {
                    int x;
                    string xs = item.GetAttribute("x");
                    if (xs.Length == 0) x = 0;
                    else if (!int.TryParse(xs, NumberStyles.Integer, CultureInfo.InvariantCulture, out x)) continue;
                    if (x >= 0 && x < count && !order.Contains(x)) order.Add(x);
                }
            }
            for (int i = 0; i < count; i++) if (!order.Contains(i)) order.Add(i);

            XmlDocument doc = top.OwnerDocument;
            XmlElement items = doc.CreateElement("items", ExcelPackage.schemaMain);
            items.SetAttribute("count", (count + 1).ToString(CultureInfo.InvariantCulture));
            foreach (int x in order)
            {
                XmlElement item = doc.CreateElement("item", ExcelPackage.schemaMain);
                if (hidden != null && hidden.Contains(x)) item.SetAttribute("h", "1");
                item.SetAttribute("x", x.ToString(CultureInfo.InvariantCulture));
                items.AppendChild(item);
            }
            XmlElement subtotal = doc.CreateElement("item", ExcelPackage.schemaMain);
            subtotal.SetAttribute("t", "default");
            items.AppendChild(subtotal);
            if (old != null) top.ReplaceChild(items, old);
            else top.InsertBefore(items, top.FirstChild); // CT_PivotField: items first
        }

        /// <summary>The numbers (x) of the hidden items of the pivot field, as the file has them.</summary>
        HashSet<int> HiddenOfXml()
        {
            var hidden = new HashSet<int>();
            foreach (XmlElement item in TopNode.SelectNodes("d:items/d:item[@h='1' or @h='true']", NameSpaceManager))
            {
                int x;
                string xs = item.GetAttribute("x");
                if (xs.Length == 0) hidden.Add(0);
                else if (int.TryParse(xs, NumberStyles.Integer, CultureInfo.InvariantCulture, out x)) hidden.Add(x);
            }
            return hidden;
        }

        void CheckItemsAllowed()
        {
            if (IsCalculated) throw new InvalidOperationException("A calculated field has no items");
            if (Grouping != null) throw new NotSupportedException("The items of a grouped field can not be set here");
        }

        /// <summary>
        /// The items of the field, made from the source when the cache has none: the shared items go to the copy of the cache
        /// of every pivot table that has it, and the pivot fields that are on an axis and have no items get them.
        /// </summary>
        List<SharedItem> EnsureItems()
        {
            CheckItemsAllowed();
            List<SharedItem> items = ReadSharedItems();
            if (items.Count == 0 || _itemsMadeHere)
            {
                items = ItemsOfSource();
                if (items.Count == 0) throw new InvalidOperationException("The source of the pivot table has no rows, so the field " + Name + " has no items");
                WriteItemsEverywhere(items); // throws for a column it can not take, before anything is written
                _itemsMadeHere = true;
            }
            else if (TopNode.SelectNodes("d:items/d:item[not(@t)]", NameSpaceManager).Count != items.Count)
            {
                WritePivotItems(items.Count, HiddenOfXml());
            }
            if (_hiddenKeys == null)
            {
                // what the file hides, by key
                _hiddenKeys = new HashSet<string>(StringComparer.Ordinal);
                foreach (int x in HiddenOfXml()) if (x < items.Count) _hiddenKeys.Add(items[x].Key);
            }
            return items;
        }

        void WriteItemsEverywhere(List<SharedItem> items)
        {
            foreach (ExcelPivotTable t in ExcelPivotTableFieldCollection.TablesOfCache(_table))
            {
                if (Index >= t.Fields.Count) continue;
                ExcelPivotTableField f = t.Fields[Index];
                WriteSharedItems((XmlElement)f._cacheFieldHelper.TopNode, items, NameSpaceManager);
                // a pivot field on an axis needs an item for each shared item
                bool onAxis = ((XmlElement)f.TopNode).GetAttribute("axis").Length > 0;
                if (ReferenceEquals(f, this) || onAxis) WritePivotItems(f, items.Count, ReferenceEquals(f, this) ? null : f.HiddenOfXml());
            }
        }

        /// <summary>Writes h="1" for the hidden keys and, for a page field, its item; asks Excel for a refresh.</summary>
        void ApplyItems(List<SharedItem> items)
        {
            var hidden = new HashSet<int>();
            for (int i = 0; i < items.Count; i++) if (_hiddenKeys != null && _hiddenKeys.Contains(items[i].Key)) hidden.Add(i);
            if (hidden.Count >= items.Count) hidden.Clear(); // one item at least is shown (all hidden can only come from cells that changed)
            WritePivotItems(items.Count, hidden);

            if (_pageFieldSettings != null)
            {
                XmlElement page = (XmlElement)_pageFieldSettings.TopNode;
                int position = -1;
                if (_pageItemSet && _pageItemKey != null && hidden.Count == 0)
                {
                    int n = 0;
                    foreach (XmlElement item in TopNode.SelectNodes("d:items/d:item[not(@t)]", NameSpaceManager))
                    {
                        int x = int.Parse(item.GetAttribute("x"), CultureInfo.InvariantCulture);
                        if (items[x].Key == _pageItemKey) position = n;
                        n++;
                    }
                }
                if (position >= 0) page.SetAttribute("item", position.ToString(CultureInfo.InvariantCulture));
                else if (_pageItemSet || hidden.Count > 0) page.RemoveAttribute("item");
                if (hidden.Count > 0) MultipleItemSelectionAllowed = true;
                else if (_pageItemSet) MultipleItemSelectionAllowed = false;
            }
            RefreshOnLoad(_table);
        }

        string[] KeysOf(List<SharedItem> items, object[] values, string argument)
        {
            if (values == null || values.Length == 0) throw new ArgumentException("No item was given", argument);
            var known = new HashSet<string>(StringComparer.Ordinal);
            foreach (SharedItem i in items) known.Add(i.Key);
            var keys = new string[values.Length];
            for (int i = 0; i < values.Length; i++)
            {
                keys[i] = KeyOf(values[i]);
                // the text of an error item (#DIV/0!), which is what ItemValues gives for it
                if (!known.Contains(keys[i]) && values[i] is string && known.Contains("e|" + (string)values[i])) keys[i] = "e|" + (string)values[i];
                if (!known.Contains(keys[i])) throw new ArgumentException("The field " + Name + " has no item " + (values[i] == null ? "(blank)" : Convert.ToString(values[i], CultureInfo.InvariantCulture)), argument);
            }
            return keys;
        }

        /// <summary>
        /// The values of the items of this field: each value of its column of the source once, in the order of the rows (a
        /// text, a number as double, a date, a boolean, the text of an error, null for the empty cells). From the cache when it
        /// has them, else from the source; nothing is written.
        /// </summary>
        public object[] ItemValues
        {
            get
            {
                CheckItemsAllowed();
                List<SharedItem> items = ReadSharedItems();
                if (items.Count == 0 || _itemsMadeHere) items = ItemsOfSource();
                var values = new object[items.Count];
                for (int i = 0; i < items.Count; i++) values[i] = items[i].Value;
                return values;
            }
        }

        /// <summary>The values of the items that are hidden. Empty when every item is shown.</summary>
        public object[] HiddenItems
        {
            get
            {
                List<SharedItem> items = ReadSharedItems();
                var list = new List<object>();
                if (_hiddenKeys != null)
                {
                    foreach (SharedItem i in items) if (_hiddenKeys.Contains(i.Key)) list.Add(i.Value);
                }
                else
                {
                    foreach (int x in HiddenOfXml()) if (x < items.Count) list.Add(items[x].Value);
                }
                return list.ToArray();
            }
        }

        /// <summary>
        /// Hides the items with these values; the other items stay as they are. A text is matched without case; a number can
        /// be given as any numeric type; null stands for the empty cells. One item at least must stay shown. Excel applies it
        /// when it opens the workbook.
        /// </summary>
        public void HideItems(params object[] values)
        {
            List<SharedItem> items = EnsureItems();
            string[] keys = KeysOf(items, values, "values");
            var hidden = new HashSet<string>(_hiddenKeys, StringComparer.Ordinal);
            foreach (string k in keys) hidden.Add(k);
            int shown = 0;
            foreach (SharedItem i in items) if (!hidden.Contains(i.Key)) shown++;
            if (shown == 0) throw new InvalidOperationException("One item of the field " + Name + " at least must be shown");
            _hiddenKeys = hidden;
            _pageItemSet = _pageFieldSettings != null;
            _pageItemKey = null;
            ApplyItems(items);
        }

        /// <summary>Shows the items with these values and hides every other item. See <see cref="HideItems(object[])"/>.</summary>
        public void ShowItems(params object[] values)
        {
            List<SharedItem> items = EnsureItems();
            var show = new HashSet<string>(KeysOf(items, values, "values"), StringComparer.Ordinal);
            _hiddenKeys = new HashSet<string>(StringComparer.Ordinal);
            foreach (SharedItem i in items) if (!show.Contains(i.Key)) _hiddenKeys.Add(i.Key);
            _pageItemSet = _pageFieldSettings != null;
            _pageItemKey = null;
            ApplyItems(items);
        }

        /// <summary>Shows every item of the field. For a page field: "(All)".</summary>
        public void ShowAllItems()
        {
            CheckItemsAllowed();
            bool pageItem = _pageFieldSettings != null && ((XmlElement)_pageFieldSettings.TopNode).HasAttribute("item");
            if (HiddenOfXml().Count == 0 && !pageItem && (_hiddenKeys == null || _hiddenKeys.Count == 0)) return; // nothing is hidden: nothing to make
            List<SharedItem> items = EnsureItems();
            _hiddenKeys = new HashSet<string>(StringComparer.Ordinal);
            _pageItemSet = _pageFieldSettings != null;
            _pageItemKey = null;
            ApplyItems(items);
        }

        /// <summary>
        /// A page field: shows the one item with this value (the drop-down of the page field shows it). The field must be a
        /// page field of the pivot table. Use <see cref="ShowAllItems"/> for "(All)" and <see cref="ShowItems(object[])"/> for
        /// several items.
        /// </summary>
        public void SetPageItem(object value)
        {
            if (_pageFieldSettings == null) throw new InvalidOperationException("The field " + Name + " is not a page field");
            List<SharedItem> items = EnsureItems();
            string key = KeysOf(items, new object[] { value }, "value")[0];
            _hiddenKeys = new HashSet<string>(StringComparer.Ordinal);
            _pageItemKey = key;
            _pageItemSet = true;
            ApplyItems(items);
        }

        /// <summary>A page field: the value of the one item it shows, or null when it shows all or several. (An item of empty cells also gives null.)</summary>
        public object PageItem
        {
            get
            {
                if (_pageFieldSettings == null) return null;
                int position;
                if (!int.TryParse(((XmlElement)_pageFieldSettings.TopNode).GetAttribute("item"), NumberStyles.Integer, CultureInfo.InvariantCulture, out position)) return null;
                List<SharedItem> items = ReadSharedItems();
                XmlNodeList list = TopNode.SelectNodes("d:items/d:item[not(@t)]", NameSpaceManager);
                if (position < 0 || position >= list.Count) return null;
                int x;
                string xs = ((XmlElement)list[position]).GetAttribute("x");
                if (xs.Length == 0) x = 0;
                else if (!int.TryParse(xs, NumberStyles.Integer, CultureInfo.InvariantCulture, out x)) return null;
                return x >= 0 && x < items.Count ? items[x].Value : null;
            }
        }

        /// <summary>
        /// Before a save: the items that were made here are made again from the source, which can have changed, and what was
        /// hidden is found again by its value. Items of a loaded workbook that were not made here are not touched.
        /// </summary>
        internal void WriteItemsOnSave()
        {
            if (!_itemsMadeHere) return;
            List<SharedItem> items = ItemsOfSource();
            if (items.Count == 0) return;
            WriteItemsEverywhere(items);
            ApplyItems(items);
        }
    }
}
