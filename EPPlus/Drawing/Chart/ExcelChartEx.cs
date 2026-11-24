// orabooster-epplus: the chart types of Excel 2016 (https://github.com/orascope/orabooster-epplus/issues/21):
// waterfall, treemap, sunburst, histogram, Pareto, box and whisker, funnel.
// Written from the chartex schema of [MS-ODRAWXML] and from a workbook in which Excel made each of these charts itself.
// (c) 2026 OraScope Co.,Ltd. GNU Lesser General Public License v3.0, as the rest of the library.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security;
using System.Text;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// The chart types that Excel 2016 added. They are made with ExcelDrawings.AddExtendedChart.
    /// </summary>
    public enum eChartExType
    {
        /// <summary>Columns that start where the one before ended: how a total is made of gains and losses</summary>
        Waterfall,
        /// <summary>Rectangles whose area is the value, grouped by the category columns</summary>
        Treemap,
        /// <summary>Rings, one per category column, whose arcs are the values</summary>
        Sunburst,
        /// <summary>How many values fall into each range of values</summary>
        Histogram,
        /// <summary>The categories from the largest to the smallest, with a line of the running share of the total</summary>
        Pareto,
        /// <summary>The quartiles, the mean and the outliers of the values of each category</summary>
        BoxWhisker,
        /// <summary>Bars of falling width, one per stage</summary>
        Funnel
    }

    /// <summary>
    /// A series of a chart of Excel 2016.
    /// </summary>
    public sealed class ExcelChartExSerie
    {
        internal string Values, Categories;

        internal ExcelChartExSerie(string values, string categories)
        {
            Values = values;
            Categories = categories;
        }

        /// <summary>The name of the series, as a text</summary>
        public string Header { get; set; }

        /// <summary>The cell with the name of the series. It is used before Header.</summary>
        public ExcelAddressBase HeaderAddress { get; set; }

        /// <summary>
        /// Waterfall: the positions (zero based) of the points that are totals. A total starts at zero and is not added to the points before it.
        /// </summary>
        public int[] Subtotals { get; set; }
    }

    /// <summary>
    /// The series of a chart of Excel 2016.
    /// </summary>
    public sealed class ExcelChartExSeries : IEnumerable<ExcelChartExSerie>
    {
        readonly ExcelChartEx _chart;
        readonly List<ExcelChartExSerie> _list = new List<ExcelChartExSerie>();

        internal ExcelChartExSeries(ExcelChartEx chart)
        {
            _chart = chart;
        }

        /// <summary>
        /// Adds a series.
        /// </summary>
        /// <param name="valuesAddress">The values, for example "B2:B20". Without a sheet name the address is on the worksheet of the chart.</param>
        /// <param name="categoriesAddress">The categories, for example "A2:A20". For a treemap or a sunburst it can have several columns, the outer group first. Null for a histogram of the values.</param>
        public ExcelChartExSerie Add(string valuesAddress, string categoriesAddress)
        {
            if (string.IsNullOrEmpty(valuesAddress)) throw new ArgumentException("No address of the values", "valuesAddress");
            var serie = new ExcelChartExSerie(_chart.FullAddress(valuesAddress), string.IsNullOrEmpty(categoriesAddress) ? null : _chart.FullAddress(categoriesAddress));
            _list.Add(serie);
            return serie;
        }

        /// <summary>
        /// Adds a series.
        /// </summary>
        /// <param name="values">The values</param>
        /// <param name="categories">The categories, or null</param>
        public ExcelChartExSerie Add(ExcelRangeBase values, ExcelRangeBase categories)
        {
            if (values == null) throw new ArgumentNullException("values");
            return Add(values.FullAddressAbsolute, categories == null ? null : categories.FullAddressAbsolute);
        }

        /// <summary>The number of series</summary>
        public int Count
        {
            get { return _list.Count; }
        }

        /// <summary>The series at a position</summary>
        public ExcelChartExSerie this[int index]
        {
            get { return _list[index]; }
        }

        /// <summary>The series in the order they were added</summary>
        public IEnumerator<ExcelChartExSerie> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
    }

    /// <summary>
    /// A chart of one of the types that Excel 2016 added. Older versions of Excel show a text box in its place.
    /// </summary>
    public sealed class ExcelChartEx : ExcelDrawing
    {
        const string NsChartEx = "http://schemas.microsoft.com/office/drawing/2014/chartex";
        const string NsChartEx1 = "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex";
        const string NsChartEx2 = "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex";
        const string RelationChartEx = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
        const string ContentTypeChartEx = "application/vnd.ms-office.chartex+xml";
        const string NamePath = "mc:AlternateContent/mc:Choice/xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name";

        internal Uri UriChart, UriStyle, UriColors;
        internal Packaging.ZipPackagePart Part, PartStyle, PartColors;

        internal ExcelChartEx(ExcelDrawings drawings, XmlNode node, eChartExType type)
            : base(drawings, WithNamespaces(drawings, node), NamePath)
        {
            ChartType = type;
            Series = new ExcelChartExSeries(this);
            DataLabels = type == eChartExType.Waterfall || type == eChartExType.Funnel || type == eChartExType.Treemap || type == eChartExType.Sunburst;
            ShowLegend = type == eChartExType.Waterfall || type == eChartExType.Treemap;

            var package = drawings.Worksheet._package.Package;
            UriChart = GetNewUri(package, "/xl/charts/chartEx{0}.xml");
            Part = package.CreatePart(UriChart, ContentTypeChartEx, drawings.Worksheet._package.Compression);
            Save();
            var relation = drawings.Part.CreateRelationship(UriHelper.GetRelativeUri(drawings.UriDrawing, UriChart), Packaging.TargetMode.Internal, RelationChartEx);

            // Excel 16 does not open a workbook in which a chart of these types has no style part and no colour part
            // (tried with its own file: with either relationship taken away the workbook is refused). The chart itself carries no
            // format, so these two parts are where its fonts, lines and colours come from.
            UriStyle = GetNewUri(package, "/xl/charts/style{0}.xml");
            PartStyle = package.CreatePart(UriStyle, "application/vnd.ms-office.chartstyle+xml", drawings.Worksheet._package.Compression);
            WriteXml(PartStyle, ExcelChartExStyle.StyleXml(type == eChartExType.Treemap));
            Part.CreateRelationship(UriHelper.GetRelativeUri(UriChart, UriStyle), Packaging.TargetMode.Internal, "http://schemas.microsoft.com/office/2011/relationships/chartStyle");
            UriColors = GetNewUri(package, "/xl/charts/colors{0}.xml");
            PartColors = package.CreatePart(UriColors, "application/vnd.ms-office.chartcolorstyle+xml", drawings.Worksheet._package.Compression);
            WriteXml(PartColors, ExcelChartExStyle.ColorsXml());
            Part.CreateRelationship(UriHelper.GetRelativeUri(UriChart, UriColors), Packaging.TargetMode.Internal, "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle");

            // The chart sits in an AlternateContent: a version of Excel that knows the namespace of the Choice shows the chart,
            // an older one shows the text box of the Fallback. Funnel charts came one release later and ask for the second namespace.
            XmlDocument doc = node.OwnerDocument;
            XmlElement alternate = doc.CreateElement("mc", "AlternateContent", ExcelPackage.schemaMarkupCompatibility);
            alternate.SetAttribute("xmlns:mc", ExcelPackage.schemaMarkupCompatibility); // declared here, so that the content below can use the prefix
            node.AppendChild(alternate);
            string prefix = type == eChartExType.Funnel ? "cx2" : "cx1";
            alternate.InnerXml =
                "<mc:Choice xmlns:" + prefix + "=\"" + (type == eChartExType.Funnel ? NsChartEx2 : NsChartEx1) + "\" Requires=\"" + prefix + "\">" +
                "<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"" + Id + "\" name=\"Chart\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>" +
                "<xdr:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm>" +
                "<a:graphic><a:graphicData uri=\"" + NsChartEx + "\"><cx:chart xmlns:cx=\"" + NsChartEx + "\" xmlns:r=\"" + ExcelPackage.schemaRelationships + "\" r:id=\"" + relation.Id + "\" /></a:graphicData></a:graphic>" +
                "</xdr:graphicFrame></mc:Choice>" +
                "<mc:Fallback><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"0\" name=\"\" /><xdr:cNvSpPr><a:spLocks noTextEdit=\"1\" /></xdr:cNvSpPr></xdr:nvSpPr>" +
                "<xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"4572000\" cy=\"2743200\" /></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst /></a:prstGeom>" +
                "<a:solidFill><a:prstClr val=\"white\" /></a:solidFill><a:ln w=\"1\"><a:solidFill><a:prstClr val=\"green\" /></a:solidFill></a:ln></xdr:spPr>" +
                "<xdr:txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\" /><a:lstStyle /><a:p><a:r><a:rPr lang=\"en-US\" sz=\"1100\" />" +
                "<a:t>This chart isn't available in your version of Excel. Editing this shape or saving this workbook into a different file format will permanently break the chart.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>";
            node.AppendChild(doc.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings));
        }

        // The paths of this drawing use the prefix mc, which the drawings of a worksheet did not need before
        static XmlNode WithNamespaces(ExcelDrawings drawings, XmlNode node)
        {
            if (!drawings.NameSpaceManager.HasNamespace("mc")) drawings.NameSpaceManager.AddNamespace("mc", ExcelPackage.schemaMarkupCompatibility);
            return node;
        }

        /// <summary>The type of the chart</summary>
        public eChartExType ChartType { get; private set; }

        /// <summary>The worksheet of the chart</summary>
        public ExcelWorksheet WorkSheet
        {
            get { return _drawings.Worksheet; }
        }

        /// <summary>The series of the chart. Most of these charts have one.</summary>
        public ExcelChartExSeries Series { get; private set; }

        /// <summary>The title above the chart. Null or empty for a chart without a title.</summary>
        public string Title { get; set; }

        /// <summary>A legend above the chart</summary>
        public bool ShowLegend { get; set; }

        /// <summary>
        /// Labels at the points: the value in a waterfall or funnel chart, the category in a treemap or sunburst chart.
        /// On for these four types when the chart is made, as in Excel.
        /// </summary>
        public bool DataLabels { get; set; }

        /// <summary>Histogram and Pareto of values: the width of a bin. Null lets Excel choose.</summary>
        public double? BinWidth { get; set; }

        /// <summary>Histogram and Pareto of values: the number of bins. Null lets Excel choose. BinWidth is used before it.</summary>
        public int? BinCount { get; set; }

        // An address as it is written in the chart: with the sheet name and absolute
        internal string FullAddress(string address)
        {
            var a = new ExcelAddressBase(address);
            if (a._fromRow < 1 || a._fromCol < 1) throw new ArgumentException("Not a cell range: " + address, "address");
            string sheet = string.IsNullOrEmpty(a.WorkSheet) ? _drawings.Worksheet.Name : a.WorkSheet;
            return ExcelCellBase.GetFullAddress(sheet, ExcelCellBase.GetAddress(a._fromRow, a._fromCol, a._toRow, a._toCol, true));
        }

        static string Esc(string text)
        {
            return SecurityElement.Escape(text ?? "");
        }

        static string Number(double value)
        {
            return value.ToString("R", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// The part of the chart, built from the properties. The elements are in the order Excel writes them.
        /// </summary>
        internal string BuildXml()
        {
            bool hierarchy = ChartType == eChartExType.Treemap || ChartType == eChartExType.Sunburst;
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<cx:chartSpace xmlns:a=\"" + ExcelPackage.schemaDrawings + "\" xmlns:r=\"" + ExcelPackage.schemaRelationships + "\" xmlns:cx=\"" + NsChartEx + "\">");

            sb.Append("<cx:chartData>");
            for (int i = 0; i < Series.Count; i++)
            {
                ExcelChartExSerie s = Series[i];
                sb.Append("<cx:data id=\"" + i.ToString(CultureInfo.InvariantCulture) + "\">");
                if (s.Categories != null) sb.Append("<cx:strDim type=\"cat\"><cx:f>" + Esc(s.Categories) + "</cx:f></cx:strDim>");
                sb.Append("<cx:numDim type=\"" + (hierarchy ? "size" : "val") + "\"><cx:f>" + Esc(s.Values) + "</cx:f></cx:numDim>");
                sb.Append("</cx:data>");
            }
            sb.Append("</cx:chartData>");

            sb.Append("<cx:chart>");
            if (!string.IsNullOrEmpty(Title))
            {
                sb.Append("<cx:title pos=\"t\" align=\"ctr\" overlay=\"0\"><cx:tx><cx:txData><cx:v>" + Esc(Title) + "</cx:v></cx:txData></cx:tx></cx:title>");
            }
            sb.Append("<cx:plotArea><cx:plotAreaRegion>");
            for (int i = 0; i < Series.Count; i++)
            {
                ExcelChartExSerie s = Series[i];
                string id = i.ToString(CultureInfo.InvariantCulture);
                sb.Append("<cx:series layoutId=\"" + LayoutId + "\">");
                AppendName(sb, s);
                if (DataLabels) AppendDataLabels(sb);
                sb.Append("<cx:dataId val=\"" + id + "\" />");
                AppendLayout(sb, s);
                if (ChartType == eChartExType.Pareto) sb.Append("<cx:axisId val=\"1\" />");
                sb.Append("</cx:series>");
                // the line of a Pareto chart is a series of its own that belongs to the columns
                if (ChartType == eChartExType.Pareto) sb.Append("<cx:series layoutId=\"paretoLine\" ownerIdx=\"" + id + "\"><cx:axisId val=\"2\" /></cx:series>");
            }
            sb.Append("</cx:plotAreaRegion>");
            AppendAxes(sb);
            sb.Append("</cx:plotArea>");
            if (ShowLegend) sb.Append("<cx:legend pos=\"t\" align=\"ctr\" overlay=\"0\" />");
            sb.Append("</cx:chart></cx:chartSpace>");
            return sb.ToString();
        }

        string LayoutId
        {
            get
            {
                switch (ChartType)
                {
                    case eChartExType.Waterfall: return "waterfall";
                    case eChartExType.Treemap: return "treemap";
                    case eChartExType.Sunburst: return "sunburst";
                    case eChartExType.BoxWhisker: return "boxWhisker";
                    case eChartExType.Funnel: return "funnel";
                    default: return "clusteredColumn"; // histogram and the columns of a Pareto chart
                }
            }
        }

        void AppendName(StringBuilder sb, ExcelChartExSerie s)
        {
            string text = s.Header;
            string formula = null;
            if (s.HeaderAddress != null)
            {
                formula = FullAddress(s.HeaderAddress.Address.IndexOf('!') >= 0 || string.IsNullOrEmpty(s.HeaderAddress.WorkSheet) ? s.HeaderAddress.Address : ExcelCellBase.GetFullAddress(s.HeaderAddress.WorkSheet, s.HeaderAddress.Address));
                var a = new ExcelAddressBase(formula);
                var ws = _drawings.Worksheet.Workbook.Worksheets[a.WorkSheet];
                if (ws != null) text = ws.Cells[a._fromRow, a._fromCol].Text;
            }
            if (formula == null && string.IsNullOrEmpty(text)) return;
            sb.Append("<cx:tx><cx:txData>");
            if (formula != null) sb.Append("<cx:f>" + Esc(formula) + "</cx:f>");
            sb.Append("<cx:v>" + Esc(text) + "</cx:v></cx:txData></cx:tx>");
        }

        void AppendDataLabels(StringBuilder sb)
        {
            switch (ChartType)
            {
                case eChartExType.Waterfall:
                    sb.Append("<cx:dataLabels pos=\"outEnd\"><cx:visibility seriesName=\"0\" categoryName=\"0\" value=\"1\" /></cx:dataLabels>");
                    break;
                case eChartExType.Treemap:
                    sb.Append("<cx:dataLabels pos=\"inEnd\"><cx:visibility seriesName=\"0\" categoryName=\"1\" value=\"0\" /></cx:dataLabels>");
                    break;
                case eChartExType.Sunburst:
                    sb.Append("<cx:dataLabels pos=\"ctr\"><cx:visibility seriesName=\"0\" categoryName=\"1\" value=\"0\" /></cx:dataLabels>");
                    break;
                default:
                    sb.Append("<cx:dataLabels><cx:visibility seriesName=\"0\" categoryName=\"0\" value=\"1\" /></cx:dataLabels>");
                    break;
            }
        }

        void AppendLayout(StringBuilder sb, ExcelChartExSerie s)
        {
            switch (ChartType)
            {
                case eChartExType.Waterfall:
                    sb.Append("<cx:layoutPr><cx:subtotals>");
                    if (s.Subtotals != null)
                    {
                        foreach (int idx in s.Subtotals) sb.Append("<cx:idx val=\"" + idx.ToString(CultureInfo.InvariantCulture) + "\" />");
                    }
                    sb.Append("</cx:subtotals></cx:layoutPr>");
                    break;
                case eChartExType.Treemap:
                    sb.Append("<cx:layoutPr><cx:parentLabelLayout val=\"overlapping\" /></cx:layoutPr>");
                    break;
                case eChartExType.BoxWhisker:
                    sb.Append("<cx:layoutPr><cx:visibility meanLine=\"0\" meanMarker=\"1\" nonoutliers=\"0\" outliers=\"1\" /><cx:statistics quartileMethod=\"exclusive\" /></cx:layoutPr>");
                    break;
                case eChartExType.Histogram:
                    AppendBinning(sb);
                    break;
                case eChartExType.Pareto:
                    // with categories the columns are the categories; without, the values are put into bins as in a histogram
                    if (s.Categories != null) sb.Append("<cx:layoutPr><cx:aggregation /></cx:layoutPr>");
                    else AppendBinning(sb);
                    break;
            }
        }

        void AppendBinning(StringBuilder sb)
        {
            sb.Append("<cx:layoutPr><cx:binning intervalClosed=\"r\"");
            if (BinWidth.HasValue) sb.Append("><cx:binSize val=\"" + Number(BinWidth.Value) + "\" /></cx:binning>");
            else if (BinCount.HasValue) sb.Append("><cx:binCount val=\"" + BinCount.Value.ToString(CultureInfo.InvariantCulture) + "\" /></cx:binning>");
            else sb.Append(" />");
            sb.Append("</cx:layoutPr>");
        }

        void AppendAxes(StringBuilder sb)
        {
            switch (ChartType)
            {
                case eChartExType.Treemap:
                case eChartExType.Sunburst:
                    break;
                case eChartExType.Funnel:
                    sb.Append("<cx:axis id=\"1\"><cx:catScaling gapWidth=\"0.06\" /><cx:tickLabels /></cx:axis>");
                    break;
                default:
                    string gap = ChartType == eChartExType.Waterfall ? "0.5" : ChartType == eChartExType.BoxWhisker ? "1" : "0";
                    sb.Append("<cx:axis id=\"0\"><cx:catScaling gapWidth=\"" + gap + "\" /><cx:tickLabels /></cx:axis>");
                    sb.Append("<cx:axis id=\"1\"><cx:valScaling /><cx:majorGridlines /><cx:tickLabels /></cx:axis>");
                    if (ChartType == eChartExType.Pareto) sb.Append("<cx:axis id=\"2\"><cx:valScaling max=\"1\" min=\"0\" /><cx:units unit=\"percentage\" /><cx:tickLabels /></cx:axis>");
                    break;
            }
        }

        /// <summary>
        /// Writes the part of the chart. Called when the worksheet is saved.
        /// </summary>
        internal void Save()
        {
            WriteXml(Part, BuildXml());
        }

        static void WriteXml(Packaging.ZipPackagePart part, string xml)
        {
            var doc = new XmlDocument();
            doc.LoadXml(xml);
            // the stream belongs to the part and is read again when the package is written: it is not closed here
            doc.Save(part.GetStream(FileMode.Create, FileAccess.Write));
        }

        internal override void DeleteMe()
        {
            Part.Package.DeletePart(UriStyle);
            Part.Package.DeletePart(UriColors);
            Part.Package.DeletePart(UriChart);
            base.DeleteMe();
        }
    }
}
