/*******************************************************************************
 * orabooster-epplus: this file is new in the fork (item B4 of issue #8, issue #46). It is not part of EPPlus 4.5 or EPPlusFree
 * and has no code from EPPlus 5 or later. Written by OraScope Co.,Ltd. 2026 from [MS-XLSX] 2.6 (the x14 namespace:
 * CT_ConditionalFormattings, CT_ConditionalFormatting, CT_CfRule, CT_DataBar, CT_Cfvo) and from files Excel made, under the
 * license of the library (LGPL-3.0).
 *
 * The part of a conditional format that Excel 2010 and later add: a rule of ECMA-376 in conditionalFormatting points with an
 * x14:id in its extLst to an x14:cfRule in the extLst of the worksheet, which has what ECMA-376 has no place for (the options
 * of a data bar).
 *
 *   <cfRule type="dataBar" priority="1"><dataBar>...</dataBar>
 *     <extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="..."><x14:id>{GUID}</x14:id></ext></extLst></cfRule>
 *   ...
 *   <extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="...">
 *     <x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="...">
 *       <x14:cfRule type="dataBar" id="{GUID}"><x14:dataBar ...><x14:cfvo type="min"/><x14:cfvo type="max"/>...</x14:dataBar></x14:cfRule>
 *       <xm:sqref>B2:B11</xm:sqref>
 *     </x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    /// <summary>How a data bar is filled (Excel 2010 and later).</summary>
    public enum eDataBarFillType
    {
        /// <summary>A gradient from the colour to white, as in Excel 2007</summary>
        Gradient,
        /// <summary>The colour</summary>
        Solid
    }

    /// <summary>Where the axis of a data bar is, between the bars of negative and of positive values (Excel 2010 and later).</summary>
    public enum eDataBarAxisPosition
    {
        /// <summary>Where Excel finds it from the values</summary>
        Automatic,
        /// <summary>In the middle of the cell</summary>
        Middle,
        /// <summary>No axis: negative values have bars in the same direction as positive ones</summary>
        None
    }

    /// <summary>The direction of a data bar (Excel 2010 and later).</summary>
    public enum eDataBarDirection
    {
        /// <summary>As the reading order of the sheet</summary>
        Context,
        /// <summary>From the left to the right</summary>
        LeftToRight,
        /// <summary>From the right to the left</summary>
        RightToLeft
    }

    /// <summary>The x14 part of the conditional formats of a worksheet.</summary>
    internal static class ExcelConditionalFormattingX14
    {
        internal const string SheetExtUri = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";
        internal const string RuleExtUri = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}";

        /// <summary>The x14:id in the extLst of a rule of ECMA-376, or null.</summary>
        internal static string GetId(XmlNode cfRule, XmlNamespaceManager ns)
        {
            if (cfRule == null) return null;
            XmlNode id = cfRule.SelectSingleNode("d:extLst/d:ext[@uri='" + RuleExtUri + "']/x14:id", ns);
            if (id == null) return null;
            string text = id.InnerText.Trim();
            return text.Length == 0 ? null : text;
        }

        static XmlNode SheetExt(ExcelWorksheet ws)
        {
            return ws.WorksheetXml.DocumentElement.SelectSingleNode("d:extLst/d:ext[@uri='" + SheetExtUri + "']", ws.NameSpaceManager);
        }

        /// <summary>The x14:cfRule with this id in the worksheet, or null.</summary>
        internal static XmlElement FindRule(ExcelWorksheet ws, string id)
        {
            if (id == null) return null;
            XmlNode ext = SheetExt(ws);
            if (ext == null) return null;
            foreach (XmlElement rule in ext.SelectNodes("x14:conditionalFormattings/x14:conditionalFormatting/x14:cfRule", ws.NameSpaceManager))
            {
                if (string.Equals(rule.GetAttribute("id"), id, StringComparison.OrdinalIgnoreCase)) return rule;
            }
            return null;
        }

        /// <summary>
        /// The x14:cfRule of a rule of ECMA-376, made when it is not there: the x14:id in the rule (a new id when it has none) and
        /// the x14:cfRule with its x14:conditionalFormatting in the worksheet.
        /// </summary>
        internal static XmlElement GetOrCreateRule(ExcelWorksheet ws, XmlElement cfRule, string type, string sqref)
        {
            XmlNamespaceManager ns = ws.NameSpaceManager;
            XmlDocument doc = cfRule.OwnerDocument;
            string id = GetId(cfRule, ns);
            XmlElement found = FindRule(ws, id);
            if (found != null) return found;

            if (id == null)
            {
                id = NewId(ws);
                // the extLst of a cfRule is its last child (CT_CfRule)
                XmlElement ruleExtLst = cfRule.SelectSingleNode("d:extLst", ns) as XmlElement;
                if (ruleExtLst == null)
                {
                    ruleExtLst = doc.CreateElement("extLst", ExcelPackage.schemaMain);
                    cfRule.AppendChild(ruleExtLst);
                }
                XmlElement ruleExt = doc.CreateElement("ext", ExcelPackage.schemaMain);
                ruleExt.SetAttribute("uri", RuleExtUri);
                ruleExt.SetAttribute("xmlns:x14", ExcelPackage.schemaMainX14);
                XmlElement idElement = doc.CreateElement("x14", "id", ExcelPackage.schemaMainX14);
                idElement.InnerText = id;
                ruleExt.AppendChild(idElement);
                ruleExtLst.AppendChild(ruleExt);
            }

            // the ext of the worksheet; the one of the conditional formats is put first, where Excel has it
            XmlNode ext = SheetExt(ws);
            if (ext == null)
            {
                XmlElement extLst = (XmlElement)ws.CreateNode("d:extLst");
                XmlElement newExt = doc.CreateElement("ext", ExcelPackage.schemaMain);
                newExt.SetAttribute("uri", SheetExtUri);
                newExt.SetAttribute("xmlns:x14", ExcelPackage.schemaMainX14);
                extLst.InsertBefore(newExt, extLst.FirstChild);
                ext = newExt;
            }
            XmlElement formattings = ext.SelectSingleNode("x14:conditionalFormattings", ns) as XmlElement;
            if (formattings == null)
            {
                formattings = doc.CreateElement("x14", "conditionalFormattings", ExcelPackage.schemaMainX14);
                ext.AppendChild(formattings);
            }
            XmlElement formatting = doc.CreateElement("x14", "conditionalFormatting", ExcelPackage.schemaMainX14);
            formatting.SetAttribute("xmlns:xm", ExcelPackage.schemaMainXm);
            formattings.AppendChild(formatting);
            XmlElement rule = doc.CreateElement("x14", "cfRule", ExcelPackage.schemaMainX14);
            rule.SetAttribute("type", type);
            rule.SetAttribute("id", id);
            formatting.AppendChild(rule);
            XmlElement sqrefElement = doc.CreateElement("xm", "sqref", ExcelPackage.schemaMainXm);
            sqrefElement.InnerText = sqref;
            formatting.AppendChild(sqrefElement);
            return rule;
        }

        /// <summary>
        /// An id for a new x14:cfRule: a GUID that no rule of the worksheet has. It is made from a number and not at random, so
        /// that the same workbook is saved with the same bytes (the samples of tools/SaveBench are compared part by part). The
        /// id only has to be the same in the rule and in its x14 part.
        /// </summary>
        static string NewId(ExcelWorksheet ws)
        {
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (XmlNode node in ws.WorksheetXml.SelectNodes("//x14:id | //x14:cfRule/@id", ws.NameSpaceManager)) used.Add(node.InnerText.Trim());
            for (long n = 1; ; n++)
            {
                string id = "{B4E1A5D0-0C1F-4F7B-9E2A-" + n.ToString("X12", CultureInfo.InvariantCulture) + "}";
                if (!used.Contains(id)) return id;
            }
        }

        /// <summary>
        /// Takes the x14:cfRule of a rule of ECMA-376 out of the worksheet, with what is empty then. Nothing happens for a rule
        /// that has none.
        /// </summary>
        internal static void RemoveRule(ExcelWorksheet ws, XmlNode cfRule)
        {
            XmlElement rule = FindRule(ws, GetId(cfRule, ws.NameSpaceManager));
            if (rule == null) return;
            XmlNode formatting = rule.ParentNode;
            formatting.RemoveChild(rule);
            if (formatting.SelectSingleNode("x14:cfRule", ws.NameSpaceManager) == null)
            {
                XmlNode formattings = formatting.ParentNode;
                formattings.RemoveChild(formatting);
                if (formattings.SelectSingleNode("x14:conditionalFormatting", ws.NameSpaceManager) == null)
                {
                    XmlNode ext = formattings.ParentNode;
                    ext.RemoveChild(formattings);
                    if (!HasElement(ext))
                    {
                        XmlNode extLst = ext.ParentNode;
                        extLst.RemoveChild(ext);
                        if (!HasElement(extLst)) extLst.ParentNode.RemoveChild(extLst);
                    }
                }
            }
        }

        static bool HasElement(XmlNode node)
        {
            foreach (XmlNode child in node.ChildNodes)
            {
                if (child is XmlElement) return true;
            }
            return false;
        }

        /// <summary>The type of an x14:cfvo for the type of a cfvo of ECMA-376 (the names are the same).</summary>
        static string CfvoType(XmlElement cfvo)
        {
            return cfvo.GetAttribute("type");
        }

        static bool SameCfvo(XmlElement cfvo, XmlElement x14Cfvo, XmlNamespaceManager ns)
        {
            string type = CfvoType(cfvo);
            string x14Type = x14Cfvo.GetAttribute("type");
            // autoMin and autoMax are of x14 only; Excel writes min and max for them in the rule of ECMA-376
            bool sameType = type == x14Type || (type == "min" && x14Type == "autoMin") || (type == "max" && x14Type == "autoMax");
            if (!sameType) return false;
            XmlNode f = x14Cfvo.SelectSingleNode("xm:f", ns);
            string x14Value = f == null ? "" : f.InnerText;
            string value = cfvo.GetAttribute("val");
            if (type == "min" || type == "max") return true;
            return value == x14Value;
        }

        static XmlElement NewX14Cfvo(XmlDocument doc, XmlElement cfvo)
        {
            XmlElement x14Cfvo = doc.CreateElement("x14", "cfvo", ExcelPackage.schemaMainX14);
            string type = CfvoType(cfvo);
            x14Cfvo.SetAttribute("type", type);
            string value = cfvo.GetAttribute("val");
            if (type != "min" && type != "max" && value.Length > 0)
            {
                XmlElement f = doc.CreateElement("xm", "f", ExcelPackage.schemaMainXm);
                f.InnerText = value;
                x14Cfvo.AppendChild(f);
            }
            return x14Cfvo;
        }

        /// <summary>
        /// Before a save: the range and the values (cfvo) of the x14 part of a data bar follow the rule, where they were changed.
        /// Where they are the same, nothing is written, so a workbook that was not changed keeps its XML.
        /// </summary>
        internal static void SyncDataBar(ExcelWorksheet ws, XmlElement cfRule, string sqref)
        {
            XmlNamespaceManager ns = ws.NameSpaceManager;
            XmlElement rule = FindRule(ws, GetId(cfRule, ns));
            if (rule == null) return;
            XmlNode sqrefNode = rule.ParentNode.SelectSingleNode("xm:sqref", ns);
            if (sqrefNode != null && sqrefNode.InnerText != sqref && rule.ParentNode.SelectNodes("x14:cfRule", ns).Count == 1)
            {
                sqrefNode.InnerText = sqref;
            }
            XmlElement dataBar = rule.SelectSingleNode("x14:dataBar", ns) as XmlElement;
            if (dataBar == null) return;
            var cfvos = new List<XmlElement>();
            foreach (XmlElement c in cfRule.SelectNodes("d:dataBar/d:cfvo", ns)) cfvos.Add(c);
            var x14Cfvos = new List<XmlElement>();
            foreach (XmlElement c in dataBar.SelectNodes("x14:cfvo", ns)) x14Cfvos.Add(c);
            if (cfvos.Count == x14Cfvos.Count)
            {
                // only the values that were changed: one that is the same stays as Excel wrote it (autoMin, autoMax)
                for (int i = 0; i < cfvos.Count; i++)
                {
                    if (!SameCfvo(cfvos[i], x14Cfvos[i], ns)) dataBar.ReplaceChild(NewX14Cfvo(dataBar.OwnerDocument, cfvos[i]), x14Cfvos[i]);
                }
                return;
            }
            XmlNode before = x14Cfvos.Count > 0 ? x14Cfvos[x14Cfvos.Count - 1].NextSibling : dataBar.FirstChild;
            foreach (XmlElement c in x14Cfvos) dataBar.RemoveChild(c);
            foreach (XmlElement c in cfvos) dataBar.InsertBefore(NewX14Cfvo(dataBar.OwnerDocument, c), before);
        }

        /// <summary>The cfvo of a new x14:dataBar, from the rule of ECMA-376.</summary>
        internal static void AddCfvos(XmlElement dataBar, XmlElement cfRule, XmlNamespaceManager ns)
        {
            foreach (XmlElement c in cfRule.SelectNodes("d:dataBar/d:cfvo", ns)) dataBar.AppendChild(NewX14Cfvo(dataBar.OwnerDocument, c));
        }

        internal static string ToRgb(System.Drawing.Color color)
        {
            return color.ToArgb().ToString("X8", CultureInfo.InvariantCulture);
        }
    }
}
