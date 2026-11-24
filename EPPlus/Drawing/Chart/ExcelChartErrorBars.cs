// orabooster-epplus: the error bars of a chart series (https://github.com/orascope/orabooster-epplus/issues/19).
// Written from ECMA-376 Part 1, 21.2 (c:errBars) and compared with charts to which Excel added error bars itself.
// (c) 2026 OraScope Co.,Ltd. GNU Lesser General Public License v3.0, as the rest of the library.
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// On which side of a point an error bar is drawn.
    /// </summary>
    public enum eErrorBarType
    {
        /// <summary>Above and below (or left and right of) the point</summary>
        Both,
        /// <summary>Below (or left of) the point</summary>
        Minus,
        /// <summary>Above (or right of) the point</summary>
        Plus
    }

    /// <summary>
    /// What the value of the error bars means.
    /// </summary>
    public enum eErrorValueType
    {
        /// <summary>The same length for every point</summary>
        FixedValue,
        /// <summary>A percentage of the value of the point</summary>
        Percentage,
        /// <summary>A number of standard deviations of the series</summary>
        StandardDeviation,
        /// <summary>The standard error of the series. The value is not used.</summary>
        StandardError
    }

    /// <summary>
    /// The error bars of a series in one direction.
    /// </summary>
    public sealed class ExcelChartErrorBars
    {
        static readonly string[] s_types = { "both", "minus", "plus" };
        static readonly string[] s_valueTypes = { "fixedVal", "percentage", "stdDev", "stdErr" };
        // the children of c:errBars in the order of the schema, then of c:spPr and a:ln
        static readonly string[] s_order =
        {
            "errDir", "errBarType", "errValType", "noEndCap", "plus", "minus", "val", "spPr", "extLst",
            "xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "ln",
            "prstDash", "custDash", "round", "bevel", "miter", "headEnd", "tailEnd"
        };

        readonly XmlElement _node;
        readonly XmlNamespaceManager _ns;
        ExcelDrawingBorder _border;

        ExcelChartErrorBars(XmlNamespaceManager ns, XmlElement node)
        {
            _ns = ns;
            _node = node;
        }

        // The error bars of a series in one direction. A bar, line or area series has one set, which is vertical for the reader
        // of the chart (a bar series writes no direction); a scatter or bubble series can have one for x and one for y.
        internal static ExcelChartErrorBars Find(ExcelChartSerie serie, bool horizontal)
        {
            foreach (XmlNode n in serie.TopNode.SelectNodes("c:errBars", serie.NameSpaceManager))
            {
                var dir = n.SelectSingleNode("c:errDir/@val", serie.NameSpaceManager);
                bool x = dir != null && dir.Value == "x";
                if (x == horizontal) return new ExcelChartErrorBars(serie.NameSpaceManager, (XmlElement)n);
            }
            return null;
        }

        internal static ExcelChartErrorBars Add(ExcelChartSerie serie, eErrorBarType type, eErrorValueType valueType, double value, bool horizontal)
        {
            string chart = serie.ChartNodeName;
            bool xy = chart == "scatterChart" || chart == "bubbleChart";
            if (!(xy || chart == "barChart" || chart == "lineChart" || chart == "areaChart"))
            {
                throw new NotSupportedException("Error bars are for bar, line, area, scatter and bubble charts, not for " + chart);
            }
            if (horizontal && !xy) throw new NotSupportedException("Horizontal error bars are for scatter and bubble charts");
            if (double.IsNaN(value) || double.IsInfinity(value) || value < 0) throw new ArgumentOutOfRangeException("value");

            ExcelChartErrorBars old = Find(serie, horizontal);
            if (old != null) old._node.ParentNode.RemoveChild(old._node);

            XmlDocument doc = serie.TopNode.OwnerDocument;
            XmlElement node = doc.CreateElement("c", "errBars", ExcelPackage.schemaChart);
            // a bar series has no direction element, as Excel writes it
            if (chart != "barChart") AppendValue(node, "errDir", horizontal ? "x" : "y");
            AppendValue(node, "errBarType", s_types[(int)type]);
            AppendValue(node, "errValType", s_valueTypes[(int)valueType]);
            AppendValue(node, "noEndCap", "0");
            if (valueType != eErrorValueType.StandardError) AppendValue(node, "val", value.ToString("R", CultureInfo.InvariantCulture));

            // the bars for x come before the bars for y
            ExcelChartErrorBars other = Find(serie, !horizontal);
            if (other != null && horizontal) serie.TopNode.InsertBefore(node, other._node);
            else if (other != null) serie.TopNode.InsertAfter(node, other._node);
            else serie.InsertInOrder(node);
            return new ExcelChartErrorBars(serie.NameSpaceManager, node);
        }

        static void AppendValue(XmlElement parent, string name, string value)
        {
            XmlElement e = parent.OwnerDocument.CreateElement("c", name, ExcelPackage.schemaChart);
            e.SetAttribute("val", value);
            parent.AppendChild(e);
        }

        string ChildValue(string name)
        {
            var n = _node.SelectSingleNode("c:" + name + "/@val", _ns);
            return n == null ? "" : n.Value;
        }

        /// <summary>
        /// True for the horizontal error bars of a scatter or bubble series.
        /// </summary>
        public bool Horizontal
        {
            get { return ChildValue("errDir") == "x"; }
        }

        /// <summary>
        /// On which side of the point the bar is drawn.
        /// </summary>
        public eErrorBarType Type
        {
            get
            {
                int i = Array.IndexOf(s_types, ChildValue("errBarType"));
                return i < 0 ? eErrorBarType.Both : (eErrorBarType)i;
            }
        }

        /// <summary>
        /// What the value means. A custom range, which this class does not make, reads as a fixed value.
        /// </summary>
        public eErrorValueType ValueType
        {
            get
            {
                int i = Array.IndexOf(s_valueTypes, ChildValue("errValType"));
                return i < 0 ? eErrorValueType.FixedValue : (eErrorValueType)i;
            }
        }

        /// <summary>
        /// The fixed value, the percentage or the number of standard deviations.
        /// </summary>
        public double Value
        {
            get
            {
                double d;
                return double.TryParse(ChildValue("val"), NumberStyles.Float, CultureInfo.InvariantCulture, out d) ? d : 0;
            }
        }

        /// <summary>
        /// False (the default) draws a short line across the end of each bar.
        /// </summary>
        public bool NoEndCap
        {
            get
            {
                string v = ChildValue("noEndCap");
                return v == "1" || v == "true";
            }
            set
            {
                XmlNode n = _node.SelectSingleNode("c:noEndCap", _ns);
                if (n != null) ((XmlElement)n).SetAttribute("val", value ? "1" : "0");
            }
        }

        /// <summary>
        /// The line of the bars: its colour, width and style.
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(_ns, _node, "c:spPr/a:ln");
                    _border.SchemaNodeOrder = s_order;
                    _border.Fill.SchemaNodeOrder = s_order;
                }
                return _border;
            }
        }
    }
}
