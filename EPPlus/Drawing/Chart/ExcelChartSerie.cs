/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		2009-12-30
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Collections;

namespace OfficeOpenXml.Drawing.Chart
{
   /// <summary>
   /// A chart serie
   /// </summary>
    public class ExcelChartSerie : XmlHelper
   {
       internal ExcelChartSeries _chartSeries;
       protected XmlNode _node;
       protected XmlNamespaceManager _ns;
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="chartSeries">Parent collection</param>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="node">Topnode</param>
       /// <param name="isPivot">Is pivotchart</param>
       internal ExcelChartSerie(ExcelChartSeries chartSeries, XmlNamespaceManager ns, XmlNode node, bool isPivot)
           : base(ns,node)
       {
           _chartSeries = chartSeries;
           _node=node;
           _ns=ns;
           // orabooster-epplus: the order of the schema (issue #19). The list was
           // idx, order, spPr, tx, marker, trendline, explosion, invertIfNegative, dLbls, cat, val, xVal, yVal, bubbleSize, bubble3D, smooth:
           // with it the line colour of a line series put c:spPr before c:tx, and the explosion of a pie slice came after the points.
           SchemaNodeOrder = SerieNodeOrder;

           if (chartSeries.Chart.ChartType == eChartType.XYScatter ||
               chartSeries.Chart.ChartType == eChartType.XYScatterLines ||
               chartSeries.Chart.ChartType == eChartType.XYScatterLinesNoMarkers ||
               chartSeries.Chart.ChartType == eChartType.XYScatterSmooth ||
               chartSeries.Chart.ChartType == eChartType.XYScatterSmoothNoMarkers ||
               chartSeries.Chart.ChartType == eChartType.Bubble ||
               chartSeries.Chart.ChartType == eChartType.Bubble3DEffect)
           {
               _seriesTopPath = "c:yVal";
               _xSeriesTopPath = "c:xVal";
           }
           else
           {
               _seriesTopPath = "c:val";
               _xSeriesTopPath = "c:cat";
           }
           _seriesPath = string.Format(_seriesPath, _seriesTopPath);

            var np = string.Format(_xSeriesPath, _xSeriesTopPath, isPivot ? "c:multiLvlStrRef" : "c:numRef");
            var sp= string.Format(_xSeriesPath, _xSeriesTopPath, isPivot ? "c:multiLvlStrRef" : "c:strRef");
            if(ExistNode(sp))
            {
                _xSeriesPath = sp;
            }
            else
            {
                _xSeriesPath = np;
            }
        }
       internal void SetID(string id)
       {
           SetXmlNodeString("c:idx/@val",id);
           SetXmlNodeString("c:order/@val", id);
       }
       const string headerPath="c:tx/c:v";
       /// <summary>
       /// Header for the serie.
       /// </summary>
       public string Header 
       {
           get
           {
                return GetXmlNodeString(headerPath);
            }
            set
            {
                Cleartx();
                SetXmlNodeString(headerPath, value);            
            }
        }

       private void Cleartx()
       {
           var n = TopNode.SelectSingleNode("c:tx", NameSpaceManager);
           if (n != null)
           {
               n.InnerXml = "";
           }
       }
       const string headerAddressPath = "c:tx/c:strRef/c:f";
        /// <summary>
       /// Header address for the serie.
       /// </summary>
       public ExcelAddressBase HeaderAddress
       {
           get
           {
               string address = GetXmlNodeString(headerAddressPath);
               if (address == "")
               {
                   return null;
               }
               else
               {
                   return new ExcelAddressBase(address);
               }
            }
            set
            {
                if ((value._fromCol != value._toCol && value._fromRow != value._toRow) || value.Addresses != null) //Single cell removed, allow row & column --> issue 15102. 
                {
                    throw (new ArgumentException("Address must be a row, column or single cell"));
                }

                Cleartx();
                SetXmlNodeString(headerAddressPath, ExcelCellBase.GetFullAddress(value.WorkSheet, value.Address));
                SetXmlNodeString("c:tx/c:strRef/c:strCache/c:ptCount/@val", "0");
            }
        }        
        string _seriesTopPath;
        string _seriesPath = "{0}/c:numRef/c:f";       
       /// <summary>
       /// Set this to a valid address or the drawing will be invalid.
       /// </summary>
       public virtual string Series
       {
           get
           {
               return GetXmlNodeString(_seriesPath);
           }
           set
           {
               CreateNode(_seriesPath,true);
               SetXmlNodeString(_seriesPath, ExcelCellBase.GetFullAddress(_chartSeries.Chart.WorkSheet.Name, value));

               if (_chartSeries.Chart.PivotTableSource != null)
               {
                    XmlNode cache = TopNode.SelectSingleNode(string.Format("{0}/c:numRef/c:numCache", _seriesTopPath), _ns);
                    if (cache != null)
                    {
                        cache.ParentNode.RemoveChild(cache);
                    }
                    SetXmlNodeString(string.Format("{0}/c:numRef/c:numCache", _seriesTopPath), "General");
               }
               
               XmlNode lit = TopNode.SelectSingleNode(string.Format("{0}/c:numLit",_seriesTopPath), _ns);
               if (lit != null)
               {
                   lit.ParentNode.RemoveChild(lit);
               }
           }

       }
       string _xSeriesTopPath;
       string _xSeriesPath = "{0}/{1}/c:f";
       /// <summary>
       /// Set an address for the horisontal labels
       /// </summary>
       public virtual string XSeries
       {
           get
           {
               return GetXmlNodeString(_xSeriesPath);
           }
           set
           {
               CreateNode(_xSeriesPath, true);
               SetXmlNodeString(_xSeriesPath, ExcelCellBase.GetFullAddress(_chartSeries.Chart.WorkSheet.Name, value));

                if(_xSeriesPath.IndexOf("c:numRef")>0)
                {
                    XmlNode cache = TopNode.SelectSingleNode(string.Format("{0}/c:numRef/c:numCache", _xSeriesTopPath), _ns);
                    if (cache != null)
                    {
                        cache.ParentNode.RemoveChild(cache);
                    }

                    XmlNode lit = TopNode.SelectSingleNode(string.Format("{0}/c:numLit", _xSeriesTopPath), _ns);
                    if (lit != null)
                    {
                        lit.ParentNode.RemoveChild(lit);
                    }
                }
                else
                {
                    XmlNode cache = TopNode.SelectSingleNode(string.Format("{0}/c:strRef/c:strCache", _xSeriesTopPath), _ns);
                    if (cache != null)
                    {
                        cache.ParentNode.RemoveChild(cache);
                    }

                    XmlNode lit = TopNode.SelectSingleNode(string.Format("{0}/c:strLit", _xSeriesTopPath), _ns);
                    if (lit != null)
                    {
                        lit.ParentNode.RemoveChild(lit);
                    }
                }
            }
       }
       ExcelChartTrendlineCollection _trendLines = null;
       /// <summary>
       /// Access to the trendline collection
       /// </summary>
        public ExcelChartTrendlineCollection TrendLines
        {
            get
            {
                if (_trendLines == null)
                {
                    _trendLines = new ExcelChartTrendlineCollection(this);
                }
                return _trendLines;
            }
        }
        // orabooster-epplus: issue #19. The children of c:ser in the order of the schema (every chart type has a part of this list,
        // in this order), then the children of c:spPr and of a:ln. The fill and the border of a series are made with this order,
        // so c:spPr goes after c:tx. It was appended after c:cat and c:val, out of the order of the schema.
        // symbol and size are the children of c:marker that come before its c:spPr; a series has no child of these names.
        internal static readonly string[] SerieNodeOrder =
        {
            "idx", "order", "tx", "symbol", "size", "spPr", "invertIfNegative", "pictureOptions", "marker", "explosion", "dPt", "dLbls", "trendline", "errBars",
            "cat", "val", "xVal", "yVal", "smooth", "shape", "bubbleSize", "bubble3D", "extLst",
            "xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "ln",
            "prstDash", "custDash", "round", "bevel", "miter", "headEnd", "tailEnd"
        };

        ExcelDrawingFill _fill = null;
        /// <summary>
        /// The fill of the series: the columns, the area, the pie slices.
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(NameSpaceManager, TopNode, "c:spPr");
                    _fill.SchemaNodeOrder = SerieNodeOrder;
                }
                return _fill;
            }
        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// The border of the series, or the line of a line series.
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(NameSpaceManager, TopNode, "c:spPr/a:ln");
                    _border.SchemaNodeOrder = SerieNodeOrder;
                    _border.Fill.SchemaNodeOrder = SerieNodeOrder;
                }
                return _border;
            }
        }

        // orabooster-epplus: issue #19. Puts a new child of c:ser where the schema has it: before the first child that comes later
        // in the order. A file with a child out of order (c:spPr at the end, from a build before this one) still gets it in a valid place.
        internal void InsertInOrder(XmlElement child)
        {
            InsertInOrder(TopNode, child);
        }

        internal static void InsertInOrder(XmlNode ser, XmlElement child)
        {
            int position = Array.IndexOf(SerieNodeOrder, child.LocalName);
            foreach (XmlNode n in ser.ChildNodes)
            {
                if (n.NodeType != XmlNodeType.Element) continue;
                int p = Array.IndexOf(SerieNodeOrder, n.LocalName);
                if (p > position)
                {
                    ser.InsertBefore(child, n);
                    return;
                }
            }
            ser.AppendChild(child);
        }

        /// <summary>
        /// The name of the chart element of the series: barChart, lineChart, pieChart ...
        /// </summary>
        internal string ChartNodeName
        {
            get { return _chartSeries.Chart.ChartNode.LocalName; }
        }

        ExcelChartDataPointCollection _dataPoints = null;
        /// <summary>
        /// Single points of the series with a format of their own: one column, one pie slice, one marker.
        /// </summary>
        public ExcelChartDataPointCollection DataPoints
        {
            get
            {
                if (_dataPoints == null) _dataPoints = new ExcelChartDataPointCollection(this);
                return _dataPoints;
            }
        }

        /// <summary>
        /// The error bars of the series, or null when it has none. For a scatter or bubble series these are the vertical ones.
        /// </summary>
        public ExcelChartErrorBars ErrorBars
        {
            get { return ExcelChartErrorBars.Find(this, false); }
        }

        /// <summary>
        /// The horizontal error bars of a scatter or bubble series, or null when it has none.
        /// </summary>
        public ExcelChartErrorBars ErrorBarsX
        {
            get { return ExcelChartErrorBars.Find(this, true); }
        }

        /// <summary>
        /// Gives the series error bars, in place of the ones it has in that direction.
        /// Bar, line, area, scatter and bubble charts have error bars; the other chart types do not.
        /// </summary>
        /// <param name="type">Above and below the point, or one of them</param>
        /// <param name="valueType">What the value means</param>
        /// <param name="value">The fixed value, the percentage or the number of standard deviations. Not used for the standard error.</param>
        /// <param name="horizontal">Horizontal error bars, for a scatter or bubble series</param>
        public ExcelChartErrorBars AddErrorBars(eErrorBarType type, eErrorValueType valueType, double value = 1, bool horizontal = false)
        {
            return ExcelChartErrorBars.Add(this, type, valueType, value, horizontal);
        }

        /// <summary>
        /// Takes the error bars off the series, in both directions.
        /// </summary>
        public void RemoveErrorBars()
        {
            foreach (XmlNode n in TopNode.SelectNodes("c:errBars", NameSpaceManager)) TopNode.RemoveChild(n);
        }
    }
}
