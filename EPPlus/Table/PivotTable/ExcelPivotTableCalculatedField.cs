/*******************************************************************************
 * orabooster-epplus: this file is new in the fork (item B2 of issue #8, issue #54). It is not part of EPPlus 4.5 or EPPlusFree
 * and has no code from EPPlus 5 or later. Written by OraScope Co.,Ltd. 2026 from ECMA-376 Part 1 (18.10.1.3 cacheField:
 * formula and databaseField; 18.10.1.69 pivotField) and from files Excel made, under the license of the library (LGPL-3.0).
 *
 * Calculated fields of a pivot table: a field whose value is a formula of other fields.
 *
 *   var margin = pivot.Fields.AddCalculatedField("Margin", "Sales-Cost");
 *   pivot.DataFields.Add(margin).Name = "Sum of Margin";
 *
 * In the file a calculated field is a cache field with a formula and no data (databaseField="0") at the end of the cache
 * fields, and a pivot field for it in every pivot table of the cache. Excel works the values out when it refreshes the
 * pivot table, which it does when the workbook is opened (refreshOnLoad).
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    public partial class ExcelPivotTableFieldCollection
    {
        /// <summary>
        /// Adds a calculated field: a field whose value is a formula of other fields of the pivot table, such as
        /// <c>Sales-Cost</c> or <c>'Unit price'*Qty</c>. Add it to the pivot table with <c>DataFields.Add</c>; a calculated field
        /// can only be a data field. Excel works its values out when the workbook is opened.
        /// </summary>
        /// <param name="name">The name of the field: not empty, and not the name of another field</param>
        /// <param name="formula">The formula, with or without the equal sign. Field names with spaces are in single quotes.</param>
        /// <returns>The field</returns>
        public ExcelPivotTableField AddCalculatedField(string name, string formula)
        {
            if (string.IsNullOrEmpty(name) || name.Trim().Length == 0) throw new ArgumentException("The name of a calculated field can not be empty", "name");
            if (this[name] != null) throw new ArgumentException("The pivot table already has a field named " + name, "name");
            formula = ExcelPivotTableField.CleanFormula(formula);
            ExcelPivotTableField.CheckFormula(_table, formula, name);

            // Every pivot table of the cache gets the cache field and a pivot field. A loaded workbook can have several pivot
            // tables on one cache, each with its own copy of the cache XML; the copy that is saved last is the one in the file.
            ExcelPivotTableField added = null;
            foreach (ExcelPivotTable table in TablesOfCache(_table))
            {
                XmlElement cacheFields = (XmlElement)table.CacheDefinition.TopNode.SelectSingleNode("d:cacheFields", table.NameSpaceManager);
                XmlElement cacheField = cacheFields.OwnerDocument.CreateElement("cacheField", ExcelPackage.schemaMain);
                cacheField.SetAttribute("name", name);
                cacheField.SetAttribute("numFmtId", "0");
                cacheField.SetAttribute("formula", formula);
                cacheField.SetAttribute("databaseField", "0");
                cacheFields.AppendChild(cacheField);
                cacheFields.SetAttribute("count", cacheFields.SelectNodes("d:cacheField", table.NameSpaceManager).Count.ToString(CultureInfo.InvariantCulture));
                // Excel works the values out on a refresh
                ((XmlElement)table.CacheDefinition.TopNode).SetAttribute("refreshOnLoad", "1");

                XmlElement pivotFields = (XmlElement)table.TopNode.SelectSingleNode("d:pivotFields", table.NameSpaceManager);
                XmlElement pivotField = pivotFields.OwnerDocument.CreateElement("pivotField", ExcelPackage.schemaMain);
                pivotField.SetAttribute("dragToRow", "0");
                pivotField.SetAttribute("dragToCol", "0");
                pivotField.SetAttribute("dragToPage", "0");
                pivotField.SetAttribute("showAll", "0");
                pivotField.SetAttribute("defaultSubtotal", "0");
                pivotFields.AppendChild(pivotField);
                int index = table.Fields.Count;
                pivotFields.SetAttribute("count", (index + 1).ToString(CultureInfo.InvariantCulture));

                var field = new ExcelPivotTableField(table.NameSpaceManager, pivotField, table, index, index);
                field.SetCacheFieldNode(cacheField);
                table.Fields.AddInternal(field);
                if (ReferenceEquals(table, _table)) added = field;
            }
            return added;
        }

        /// <summary>The pivot tables of the workbook that have the same cache part as this one (this one among them).</summary>
        internal static List<ExcelPivotTable> TablesOfCache(ExcelPivotTable pivotTable)
        {
            var list = new List<ExcelPivotTable>();
            string uri = pivotTable.CacheDefinition.CacheDefinitionUri.OriginalString;
            foreach (ExcelWorksheet ws in pivotTable.WorkSheet.Workbook.Worksheets)
            {
                if (ws is ExcelChartsheet) continue;
                foreach (ExcelPivotTable t in ws.PivotTables)
                {
                    if (string.Equals(t.CacheDefinition.CacheDefinitionUri.OriginalString, uri, StringComparison.OrdinalIgnoreCase)) list.Add(t);
                }
            }
            if (!list.Contains(pivotTable)) list.Add(pivotTable); // a pivot table that is being made is not in the collection yet
            return list;
        }
    }

    public partial class ExcelPivotTableField
    {
        internal static string CleanFormula(string formula)
        {
            if (formula == null) throw new ArgumentNullException("formula");
            formula = formula.Trim();
            if (formula.StartsWith("=", StringComparison.Ordinal)) formula = formula.Substring(1).Trim();
            if (formula.Length == 0) throw new ArgumentException("The formula of a calculated field can not be empty", "formula");
            return formula;
        }

        /// <summary>
        /// Throws if the formula names a field the pivot table does not have, or the calculated field itself: Excel does not
        /// open a workbook with such a formula. A name is a word (letters, digits, underscore, dot) that is not followed by a
        /// bracket (that is a function), or a text in single quotes; numbers, TRUE and FALSE and texts in double quotes are not names.
        /// </summary>
        internal static void CheckFormula(ExcelPivotTable table, string formula, string ownName)
        {
            int i = 0;
            while (i < formula.Length)
            {
                char c = formula[i];
                if (c == '"')
                {
                    i++;
                    while (i < formula.Length)
                    {
                        if (formula[i] == '"') { if (i + 1 < formula.Length && formula[i + 1] == '"') { i += 2; continue; } break; }
                        i++;
                    }
                    i++;
                }
                else if (c == '\'')
                {
                    var name = new System.Text.StringBuilder();
                    i++;
                    bool closed = false;
                    while (i < formula.Length)
                    {
                        if (formula[i] == '\'')
                        {
                            if (i + 1 < formula.Length && formula[i + 1] == '\'') { name.Append('\''); i += 2; continue; }
                            closed = true;
                            break;
                        }
                        name.Append(formula[i++]);
                    }
                    if (!closed) throw new ArgumentException("A field name in single quotes is not closed: " + formula, "formula");
                    i++;
                    CheckFieldName(table, name.ToString(), ownName);
                }
                else if (char.IsDigit(c) || (c == '.' && i + 1 < formula.Length && char.IsDigit(formula[i + 1])))
                {
                    while (i < formula.Length && (char.IsDigit(formula[i]) || formula[i] == '.')) i++;
                    if (i < formula.Length && (formula[i] == 'E' || formula[i] == 'e'))
                    {
                        int j = i + 1;
                        if (j < formula.Length && (formula[j] == '+' || formula[j] == '-')) j++;
                        if (j < formula.Length && char.IsDigit(formula[j]))
                        {
                            i = j;
                            while (i < formula.Length && char.IsDigit(formula[i])) i++;
                        }
                    }
                }
                else if (char.IsLetter(c) || c == '_')
                {
                    int start = i;
                    while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] == '_' || formula[i] == '.')) i++;
                    string word = formula.Substring(start, i - start);
                    int j = i;
                    while (j < formula.Length && formula[j] == ' ') j++;
                    bool function = j < formula.Length && formula[j] == '(';
                    if (function || word.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || word.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) continue;
                    CheckFieldName(table, word, ownName);
                }
                else
                {
                    i++;
                }
            }
        }

        static void CheckFieldName(ExcelPivotTable table, string name, string ownName)
        {
            if (string.Equals(name, ownName, StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("The formula of the calculated field " + ownName + " names the field itself", "formula");
            if (table.Fields[name] == null) throw new ArgumentException("The formula names a field the pivot table does not have: " + name, "formula");
        }

        /// <summary>True if the field is a calculated field: its value is a formula of other fields (see <see cref="Formula"/>).</summary>
        public bool IsCalculated
        {
            get { return _cacheFieldHelper != null && ((XmlElement)_cacheFieldHelper.TopNode).GetAttribute("formula").Length > 0; }
        }

        /// <summary>
        /// The formula of a calculated field, without the equal sign; null for a field of the source data. Setting it is for a
        /// calculated field only (make one with <see cref="ExcelPivotTableFieldCollection.AddCalculatedField(string, string)"/>).
        /// </summary>
        public string Formula
        {
            get { return IsCalculated ? ((XmlElement)_cacheFieldHelper.TopNode).GetAttribute("formula") : null; }
            set
            {
                if (!IsCalculated) throw new InvalidOperationException("Only a calculated field has a formula");
                string formula = CleanFormula(value);
                // in the copy of the cache of every pivot table that has it
                string name = ((XmlElement)_cacheFieldHelper.TopNode).GetAttribute("name");
                CheckFormula(_table, formula, name);
                foreach (ExcelWorksheet ws in _table.WorkSheet.Workbook.Worksheets)
                {
                    if (ws is ExcelChartsheet) continue;
                    foreach (ExcelPivotTable t in ws.PivotTables)
                    {
                        if (!string.Equals(t.CacheDefinition.CacheDefinitionUri.OriginalString, _table.CacheDefinition.CacheDefinitionUri.OriginalString, StringComparison.OrdinalIgnoreCase)) continue;
                        foreach (XmlElement cacheField in t.CacheDefinition.TopNode.SelectNodes("d:cacheFields/d:cacheField", t.NameSpaceManager))
                        {
                            if (cacheField.GetAttribute("name") == name && cacheField.GetAttribute("formula").Length > 0)
                            {
                                cacheField.SetAttribute("formula", formula);
                                ((XmlElement)t.CacheDefinition.TopNode).SetAttribute("refreshOnLoad", "1");
                            }
                        }
                    }
                }
                ((XmlElement)_cacheFieldHelper.TopNode).SetAttribute("formula", formula);
            }
        }
    }
}
