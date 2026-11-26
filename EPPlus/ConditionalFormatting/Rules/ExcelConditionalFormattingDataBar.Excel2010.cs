/*******************************************************************************
 * orabooster-epplus: this file is new in the fork (item B4 of issue #8, issue #46). It is not part of EPPlus 4.5 or EPPlusFree
 * and has no code from EPPlus 5 or later. Written by OraScope Co.,Ltd. 2026 from [MS-XLSX] 2.6 (CT_DataBar of the x14
 * namespace) and from files Excel made, under the license of the library (LGPL-3.0).
 *
 * The options of a data bar that Excel 2010 and later have:
 *
 *   var bar = sheet.ConditionalFormatting.AddDatabar(range, Color.FromArgb(0x00, 0x70, 0xC0));
 *   bar.FillType = eDataBarFillType.Solid;
 *   bar.ShowBorder = true;
 *   bar.NegativeFillColor = Color.Red;
 *   bar.AxisPosition = eDataBarAxisPosition.Middle;
 *
 * They are in the x14:dataBar of the x14 part of the rule (ExcelConditionalFormattingX14), which is made when an option is set
 * first. A data bar that sets none has no x14 part and is the data bar of Excel 2007, as before.
 *******************************************************************************/
using System;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    public partial class ExcelConditionalFormattingDataBar
    {
        // the children of x14:dataBar after its two x14:cfvo, in the order of CT_DataBar
        static readonly string[] _colorOrder = { "fillColor", "borderColor", "negativeFillColor", "negativeBorderColor", "axisColor" };

        XmlElement X14DataBar(bool create)
        {
            ExcelWorksheet ws = RuleWorksheet;
            XmlElement rule = create
                ? ExcelConditionalFormattingX14.GetOrCreateRule(ws, (XmlElement)Node, "dataBar", Address.AddressSpaceSeparated)
                : ExcelConditionalFormattingX14.FindRule(ws, ExcelConditionalFormattingX14.GetId(Node, NameSpaceManager));
            if (rule == null) return null;
            XmlElement dataBar = rule.SelectSingleNode("x14:dataBar", NameSpaceManager) as XmlElement;
            if (dataBar == null && create)
            {
                dataBar = rule.OwnerDocument.CreateElement("x14", "dataBar", ExcelPackage.schemaMainX14);
                // As the bar of Excel 2007 looks until another option is set: negative bars as the positive ones (the schema has
                // false as the default) and no axis (Excel 2016 shows a bar with no x14 part with AxisPosition none). The other
                // defaults of the schema are already the look of Excel 2007: gradient, no border, 10 to 90 percent.
                dataBar.SetAttribute("negativeBarColorSameAsPositive", "1");
                dataBar.SetAttribute("axisPosition", "none");
                rule.AppendChild(dataBar);
                ExcelConditionalFormattingX14.AddCfvos(dataBar, (XmlElement)Node, NameSpaceManager);
            }
            return dataBar;
        }

        string GetX14Attribute(string name)
        {
            XmlElement dataBar = X14DataBar(false);
            return dataBar == null ? "" : dataBar.GetAttribute(name);
        }

        void SetX14Attribute(string name, string value)
        {
            XmlElement dataBar = X14DataBar(true);
            if (value == null) dataBar.RemoveAttribute(name);
            else dataBar.SetAttribute(name, value);
        }

        Color? GetX14Color(string name)
        {
            XmlElement dataBar = X14DataBar(false);
            XmlElement color = dataBar == null ? null : dataBar.SelectSingleNode("x14:" + name, NameSpaceManager) as XmlElement;
            if (color == null) return null;
            string rgb = color.GetAttribute("rgb");
            int argb;
            if (rgb.Length > 0 && int.TryParse(rgb, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out argb)) return Color.FromArgb(argb);
            return null; // a theme or an indexed colour: there is no property for it, and it is kept as it is
        }

        void SetX14Color(string name, Color? value)
        {
            XmlElement dataBar = X14DataBar(true);
            XmlNode old = dataBar.SelectSingleNode("x14:" + name, NameSpaceManager);
            if (old != null) dataBar.RemoveChild(old);
            if (value == null) return;
            XmlElement color = dataBar.OwnerDocument.CreateElement("x14", name, ExcelPackage.schemaMainX14);
            color.SetAttribute("rgb", ExcelConditionalFormattingX14.ToRgb(value.Value));
            int pos = Array.IndexOf(_colorOrder, name);
            XmlNode before = null;
            foreach (XmlNode child in dataBar.ChildNodes)
            {
                if (child.NamespaceURI == ExcelPackage.schemaMainX14 && Array.IndexOf(_colorOrder, child.LocalName) > pos) { before = child; break; }
            }
            dataBar.InsertBefore(color, before);
        }

        static bool IsTrue(string value)
        {
            return value == "1" || value.Equals("true", StringComparison.OrdinalIgnoreCase);
        }

        static bool IsFalse(string value)
        {
            return value == "0" || value.Equals("false", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>True if the rule has the part of Excel 2010 and later, where the options below are.</summary>
        public bool HasExcel2010Options
        {
            get { return X14DataBar(false) != null; }
        }

        /// <summary>How the bar is filled: a gradient (as in Excel 2007, the default) or the colour.</summary>
        public eDataBarFillType FillType
        {
            get { return IsFalse(GetX14Attribute("gradient")) ? eDataBarFillType.Solid : eDataBarFillType.Gradient; }
            set { SetX14Attribute("gradient", value == eDataBarFillType.Solid ? "0" : null); }
        }

        /// <summary>
        /// True if the bar has a border. Its colour is <see cref="BorderColor"/>; when that is not set, the colour of the bar
        /// is used.
        /// </summary>
        public bool ShowBorder
        {
            get { return IsTrue(GetX14Attribute("border")); }
            set
            {
                SetX14Attribute("border", value ? "1" : null);
                if (value && BorderColor == null) BorderColor = Color;
            }
        }

        /// <summary>The colour of the border, or null. See <see cref="ShowBorder"/>.</summary>
        public Color? BorderColor
        {
            get { return GetX14Color("borderColor"); }
            set { SetX14Color("borderColor", value); }
        }

        /// <summary>The colour of the bars of negative values; null for the colour of the bar.</summary>
        public Color? NegativeFillColor
        {
            get
            {
                if (IsTrue(GetX14Attribute("negativeBarColorSameAsPositive"))) return null;
                return GetX14Color("negativeFillColor");
            }
            set
            {
                SetX14Color("negativeFillColor", value);
                SetX14Attribute("negativeBarColorSameAsPositive", value == null ? "1" : null);
            }
        }

        /// <summary>The colour of the border of the bars of negative values; null for the colour of the border of the bar.</summary>
        public Color? NegativeBorderColor
        {
            get
            {
                if (!IsFalse(GetX14Attribute("negativeBarBorderColorSameAsPositive"))) return null; // the default is true
                return GetX14Color("negativeBorderColor");
            }
            set
            {
                SetX14Color("negativeBorderColor", value);
                SetX14Attribute("negativeBarBorderColorSameAsPositive", value == null ? null : "0");
            }
        }

        /// <summary>
        /// Where the axis between the bars of negative and of positive values is. A bar with no part of Excel 2010 has none; in
        /// the part the default is automatic.
        /// </summary>
        public eDataBarAxisPosition AxisPosition
        {
            get
            {
                if (X14DataBar(false) == null) return eDataBarAxisPosition.None;
                switch (GetX14Attribute("axisPosition"))
                {
                    case "middle": return eDataBarAxisPosition.Middle;
                    case "none": return eDataBarAxisPosition.None;
                    default: return eDataBarAxisPosition.Automatic;
                }
            }
            set
            {
                SetX14Attribute("axisPosition", value == eDataBarAxisPosition.Middle ? "middle" : value == eDataBarAxisPosition.None ? "none" : null);
            }
        }

        /// <summary>The colour of the axis, or null for the colour Excel takes (black).</summary>
        public Color? AxisColor
        {
            get { return GetX14Color("axisColor"); }
            set { SetX14Color("axisColor", value); }
        }

        /// <summary>The direction of the bar. The default is the reading order of the sheet.</summary>
        public eDataBarDirection Direction
        {
            get
            {
                switch (GetX14Attribute("direction"))
                {
                    case "leftToRight": return eDataBarDirection.LeftToRight;
                    case "rightToLeft": return eDataBarDirection.RightToLeft;
                    default: return eDataBarDirection.Context;
                }
            }
            set
            {
                SetX14Attribute("direction", value == eDataBarDirection.LeftToRight ? "leftToRight" : value == eDataBarDirection.RightToLeft ? "rightToLeft" : null);
            }
        }

        /// <summary>
        /// The length of the shortest bar, in percent of the width of the cell: 0 to 100. The default is 10, as in Excel 2007;
        /// Excel 2010 and later make new bars with 0.
        /// </summary>
        public int MinLength
        {
            get { return GetLength("minLength", 10); }
            set { SetLength("minLength", value); }
        }

        /// <summary>
        /// The length of the longest bar, in percent of the width of the cell: 0 to 100. The default is 90, as in Excel 2007;
        /// Excel 2010 and later make new bars with 100.
        /// </summary>
        public int MaxLength
        {
            get { return GetLength("maxLength", 90); }
            set { SetLength("maxLength", value); }
        }

        int GetLength(string name, int byDefault)
        {
            int length;
            return int.TryParse(GetX14Attribute(name), NumberStyles.Integer, CultureInfo.InvariantCulture, out length) ? length : byDefault;
        }

        void SetLength(string name, int value)
        {
            if (value < 0 || value > 100) throw new ArgumentOutOfRangeException(name, "The length of a data bar is 0 to 100 percent");
            SetX14Attribute(name, value.ToString(CultureInfo.InvariantCulture));
        }
    }
}
