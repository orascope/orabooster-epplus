using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;
using System.Drawing;

namespace OfficeOpenXml.Style.Dxf
{
    public class ExcelDxfStyleConditionalFormatting : DxfStyleBase<ExcelDxfStyleConditionalFormatting>
    {
        XmlHelperInstance _helper;
        internal ExcelDxfStyleConditionalFormatting(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles) : base(styles)
        {
            NumberFormat = new ExcelDxfNumberFormat(_styles);
            Font = new ExcelDxfFontBase(_styles);
            Border = new ExcelDxfBorderBase(_styles);
            Fill = new ExcelDxfFill(_styles);
            if (topNode != null)
            {
                _helper = new XmlHelperInstance(nameSpaceManager, topNode);
                // orabooster-epplus (issue #40): what the dxf does not have is "not set" here, so that a format that is loaded,
                // changed and written again says what the file said and no more. Before:
                //  - a dxf with no numFmt was read as the format "General", and written with <numFmt numFmtId="0" formatCode="General"/>
                //  - <b/>, <i/> and <strike/> with no val (true, and the way Excel writes them) were read as not set; <u/> too
                //  - a patternFill with no patternType was read as None and written as patternType="none": the fill was gone.
                //    Excel writes the fill of a dxf with only a background colour
                //  - a border with no left, right, top or bottom was read as None and written with style="none" on all four
                if (topNode.SelectSingleNode("d:numFmt", nameSpaceManager) != null)
                {
                    NumberFormat.NumFmtID = _helper.GetXmlNodeInt("d:numFmt/@numFmtId");
                    NumberFormat.Format = _helper.GetXmlNodeString("d:numFmt/@formatCode");
                    if (NumberFormat.NumFmtID < 164 && string.IsNullOrEmpty(NumberFormat.Format))
                    {
                        NumberFormat.Format = ExcelNumberFormat.GetFromBuildInFromID(NumberFormat.NumFmtID);
                    }
                }

                Font.Bold = ReadFlag(topNode, nameSpaceManager, "d:font/d:b");
                Font.Italic = ReadFlag(topNode, nameSpaceManager, "d:font/d:i");
                Font.Strike = ReadFlag(topNode, nameSpaceManager, "d:font/d:strike");
                XmlElement underline = topNode.SelectSingleNode("d:font/d:u", nameSpaceManager) as XmlElement;
                if (underline != null)
                {
                    string u = underline.GetAttribute("val");
                    Font.Underline = u == "" ? ExcelUnderLineType.Single : u == "none" ? ExcelUnderLineType.None : GetUnderLineEnum(u);
                }
                Font.Color = GetColor(_helper, "d:font/d:color");

                Border.Left = GetBorderItem(_helper, "d:border/d:left");
                Border.Right = GetBorderItem(_helper, "d:border/d:right");
                Border.Bottom = GetBorderItem(_helper, "d:border/d:bottom");
                Border.Top = GetBorderItem(_helper, "d:border/d:top");

                string patternType = _helper.GetXmlNodeString("d:fill/d:patternFill/@patternType");
                Fill.PatternType = patternType == "" ? (ExcelFillStyle?)null : GetPatternTypeEnum(patternType);
                Fill.BackgroundColor = GetColor(_helper, "d:fill/d:patternFill/d:bgColor/");
                Fill.PatternColor = GetColor(_helper, "d:fill/d:patternFill/d:fgColor/");
            }
            else
            {
                _helper = new XmlHelperInstance(nameSpaceManager);
            }
            _helper.SchemaNodeOrder = new string[] { "font", "numFmt", "fill", "border" };
        }
        // orabooster-epplus (issue #40): an on/off element of a font (b, i, strike): null when it is not there, else its val, true with no val
        private static bool? ReadFlag(XmlNode topNode, XmlNamespaceManager nameSpaceManager, string path)
        {
            XmlElement node = topNode.SelectSingleNode(path, nameSpaceManager) as XmlElement;
            if (node == null) return null;
            string val = node.GetAttribute("val");
            if (val == "") return true;
            return val == "1" || val.Equals("true", StringComparison.OrdinalIgnoreCase);
        }
        private ExcelDxfBorderItem GetBorderItem(XmlHelperInstance helper, string path)
        {
            ExcelDxfBorderItem bi = new ExcelDxfBorderItem(_styles);
            string style = helper.GetXmlNodeString(path + "/@style");
            bi.Style = style == "" ? (ExcelBorderStyle?)null : GetBorderStyleEnum(style); // orabooster-epplus (issue #40): no style is "not set", not None
            bi.Color = GetColor(helper, path+"/d:color");
            return bi;
        }
        private ExcelBorderStyle GetBorderStyleEnum(string style)
        {
            if (style == "") return ExcelBorderStyle.None;
            string sInStyle = style.Substring(0, 1).ToUpper(CultureInfo.InvariantCulture) + style.Substring(1, style.Length - 1);
            try
            {
                return (ExcelBorderStyle)Enum.Parse(typeof(ExcelBorderStyle), sInStyle);
            }
            catch
            {
                return ExcelBorderStyle.None;
            }

        }
        private ExcelFillStyle GetPatternTypeEnum(string patternType)
        {
            if (patternType == "") return ExcelFillStyle.None;
            patternType = patternType.Substring(0, 1).ToUpper(CultureInfo.InvariantCulture) + patternType.Substring(1, patternType.Length - 1);
            try
            {
                return (ExcelFillStyle)Enum.Parse(typeof(ExcelFillStyle), patternType);
            }
            catch
            {
                return ExcelFillStyle.None;
            }
        }
        private ExcelDxfColor GetColor(XmlHelperInstance helper, string path)
        {            
            ExcelDxfColor ret = new ExcelDxfColor(_styles);
            ret.Theme = helper.GetXmlNodeIntNull(path + "/@theme");
            ret.Index = helper.GetXmlNodeIntNull(path + "/@indexed");
            string rgb=helper.GetXmlNodeString(path + "/@rgb");
            if(rgb!="")
            {
                ret.Color = Color.FromArgb( int.Parse(rgb.Substring(0, 2), System.Globalization.NumberStyles.AllowHexSpecifier),
                                            int.Parse(rgb.Substring(2, 2), System.Globalization.NumberStyles.AllowHexSpecifier),
                                            int.Parse(rgb.Substring(4, 2), System.Globalization.NumberStyles.AllowHexSpecifier),
                                            int.Parse(rgb.Substring(6, 2), System.Globalization.NumberStyles.AllowHexSpecifier));
            }
            ret.Auto = helper.GetXmlNodeBoolNullable(path + "/@auto");
            ret.Tint = helper.GetXmlNodeDoubleNull(path + "/@tint");
            return ret;
        }
        private ExcelUnderLineType? GetUnderLineEnum(string value)
        {
            switch(value.ToLower(CultureInfo.InvariantCulture))
            {
                case "single":
                    return ExcelUnderLineType.Single;
                case "double":
                    return ExcelUnderLineType.Double;
                case "singleaccounting":
                    return ExcelUnderLineType.SingleAccounting;
                case "doubleaccounting":
                    return ExcelUnderLineType.DoubleAccounting;
                default:
                    return null;
            }
        }
        internal int DxfId { get; set; }
        public ExcelDxfFontBase Font { get; set; }
        public ExcelDxfNumberFormat NumberFormat { get; set; }
        public ExcelDxfFill Fill { get; set; }
        public ExcelDxfBorderBase Border { get; set; }
        protected internal override string Id
        {
            get
            {
                return NumberFormat.Id + Font.Id + Border.Id + Fill.Id +
                    (AllowChange ? "" : DxfId.ToString());//If allowchange is false we add the dxfID to ensure it's not used when conditional formatting is updated);
            }
        }
        protected internal override ExcelDxfStyleConditionalFormatting Clone()
        {
 	       var s=new ExcelDxfStyleConditionalFormatting(_helper.NameSpaceManager, null, _styles);
           s.Font = Font.Clone();
           s.NumberFormat = NumberFormat.Clone();
           s.Fill = Fill.Clone();
           s.Border = Border.Clone();
           return s;
        }

        protected internal override void CreateNodes(XmlHelper helper, string path)
        {
            if(Font.HasValue) Font.CreateNodes(helper, "d:font");
            if (NumberFormat.HasValue) NumberFormat.CreateNodes(helper, "d:numFmt");            
            if (Fill.HasValue) Fill.CreateNodes(helper, "d:fill");
            if (Border.HasValue) Border.CreateNodes(helper, "d:border");
        }
        protected internal override bool HasValue
        {
            get { return Font.HasValue || NumberFormat.HasValue || Fill.HasValue || Border.HasValue; }
        }
    }
}
