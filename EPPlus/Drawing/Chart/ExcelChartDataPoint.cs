// orabooster-epplus: the format of single data points of a chart series (https://github.com/orascope/orabooster-epplus/issues/19).
// Written from ECMA-376 Part 1, 21.2 (c:dPt) and compared with charts in which Excel coloured a point itself.
// (c) 2026 OraScope Co.,Ltd. GNU Lesser General Public License v3.0, as the rest of the library.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// One point of a series with a format of its own.
    /// </summary>
    public sealed class ExcelChartDataPoint
    {
        // the children of c:dPt in the order of the schema, then of c:marker, c:spPr and a:ln
        static readonly string[] s_order =
        {
            "idx", "invertIfNegative", "marker", "bubble3D", "explosion", "spPr", "pictureOptions", "extLst",
            "symbol", "size",
            "xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "ln",
            "prstDash", "custDash", "round", "bevel", "miter", "headEnd", "tailEnd"
        };

        internal readonly XmlElement Node;
        readonly XmlNamespaceManager _ns;
        ExcelDrawingFill _fill, _markerFill;
        ExcelDrawingBorder _border, _markerBorder;

        internal ExcelChartDataPoint(XmlNamespaceManager ns, XmlElement node)
        {
            _ns = ns;
            Node = node;
        }

        /// <summary>
        /// The position of the point in the series, zero based.
        /// </summary>
        public int Index
        {
            get
            {
                var idx = Node.SelectSingleNode("c:idx/@val", _ns);
                int i;
                return idx != null && int.TryParse(idx.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out i) ? i : -1;
            }
        }

        /// <summary>
        /// The fill of the point: its column, its pie slice, its bubble.
        /// For a point of a line or scatter series use MarkerFill; this would be the fill under the line.
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(_ns, Node, "c:spPr");
                    _fill.SchemaNodeOrder = s_order;
                }
                return _fill;
            }
        }

        /// <summary>
        /// The border of the point. In a line series this is the piece of the line that ends at the point.
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(_ns, Node, "c:spPr/a:ln");
                    _border.SchemaNodeOrder = s_order;
                    _border.Fill.SchemaNodeOrder = s_order;
                }
                return _border;
            }
        }

        /// <summary>
        /// The fill of the marker of the point, in a line, scatter or radar series.
        /// </summary>
        public ExcelDrawingFill MarkerFill
        {
            get
            {
                if (_markerFill == null)
                {
                    _markerFill = new ExcelDrawingFill(_ns, Node, "c:marker/c:spPr");
                    _markerFill.SchemaNodeOrder = s_order;
                }
                return _markerFill;
            }
        }

        /// <summary>
        /// The outline of the marker of the point, in a line, scatter or radar series.
        /// </summary>
        public ExcelDrawingBorder MarkerBorder
        {
            get
            {
                if (_markerBorder == null)
                {
                    _markerBorder = new ExcelDrawingBorder(_ns, Node, "c:marker/c:spPr/a:ln");
                    _markerBorder.SchemaNodeOrder = s_order;
                    _markerBorder.Fill.SchemaNodeOrder = s_order;
                }
                return _markerBorder;
            }
        }
    }

    /// <summary>
    /// The points of a series that have a format of their own.
    /// </summary>
    public sealed class ExcelChartDataPointCollection : IEnumerable<ExcelChartDataPoint>
    {
        readonly ExcelChartSerie _serie;
        readonly List<ExcelChartDataPoint> _list = new List<ExcelChartDataPoint>();

        internal ExcelChartDataPointCollection(ExcelChartSerie serie)
        {
            _serie = serie;
            foreach (XmlNode n in serie.TopNode.SelectNodes("c:dPt", serie.NameSpaceManager))
            {
                var e = n as XmlElement;
                if (e != null) _list.Add(new ExcelChartDataPoint(serie.NameSpaceManager, e));
            }
        }

        /// <summary>
        /// The point at a position of the series, zero based. It is made when the series has no format for that point yet.
        /// A surface chart has no points of its own.
        /// </summary>
        /// <param name="index">The position of the point in the series, zero based</param>
        public ExcelChartDataPoint Add(int index)
        {
            if (index < 0) throw new ArgumentOutOfRangeException("index");
            foreach (var p in _list)
            {
                if (p.Index == index) return p;
            }
            string chart = _serie.ChartNodeName;
            if (chart == "surfaceChart" || chart == "surface3DChart")
            {
                throw new NotSupportedException("A surface chart has no data points with a format of their own");
            }

            XmlDocument doc = _serie.TopNode.OwnerDocument;
            XmlElement node = doc.CreateElement("c", "dPt", ExcelPackage.schemaChart);
            // as Excel writes a point: the index, invertIfNegative in a bar or bubble series, and bubble3D
            AppendValue(node, "idx", index.ToString(CultureInfo.InvariantCulture));
            if (chart == "barChart" || chart == "bar3DChart" || chart == "bubbleChart") AppendValue(node, "invertIfNegative", "0");
            AppendValue(node, "bubble3D", "0");

            // the points are kept in the order of their index
            XmlNode before = null;
            foreach (var p in _list)
            {
                if (p.Index > index && p.Node.ParentNode == _serie.TopNode) { before = p.Node; break; }
            }
            if (before != null) _serie.TopNode.InsertBefore(node, before);
            else if (_list.Count > 0 && _list[_list.Count - 1].Node.ParentNode == _serie.TopNode) _serie.TopNode.InsertAfter(node, _list[_list.Count - 1].Node);
            else _serie.InsertInOrder(node);

            var point = new ExcelChartDataPoint(_serie.NameSpaceManager, node);
            int at = _list.FindIndex(p => p.Index > index);
            if (at < 0) _list.Add(point); else _list.Insert(at, point);
            return point;
        }

        static void AppendValue(XmlElement parent, string name, string value)
        {
            XmlElement e = parent.OwnerDocument.CreateElement("c", name, ExcelPackage.schemaChart);
            e.SetAttribute("val", value);
            parent.AppendChild(e);
        }

        /// <summary>
        /// The point at a position of the series, or null when the series has no format for it.
        /// </summary>
        /// <param name="index">The position of the point in the series, zero based</param>
        public ExcelChartDataPoint this[int index]
        {
            get
            {
                foreach (var p in _list)
                {
                    if (p.Index == index) return p;
                }
                return null;
            }
        }

        /// <summary>
        /// Takes the format off a point. It is drawn as the rest of the series again.
        /// </summary>
        /// <param name="index">The position of the point in the series, zero based</param>
        public bool Remove(int index)
        {
            ExcelChartDataPoint p = this[index];
            if (p == null) return false;
            if (p.Node.ParentNode != null) p.Node.ParentNode.RemoveChild(p.Node);
            _list.Remove(p);
            return true;
        }

        /// <summary>The number of points with a format of their own</summary>
        public int Count
        {
            get { return _list.Count; }
        }

        /// <summary>The points, in the order of their position</summary>
        public IEnumerator<ExcelChartDataPoint> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
    }
}
