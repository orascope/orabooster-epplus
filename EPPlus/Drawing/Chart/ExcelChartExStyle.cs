// orabooster-epplus: the style part and the colour part of a chart of Excel 2016 (https://github.com/orascope/orabooster-epplus/issues/21).
// Excel does not open a workbook in which such a chart has no style part and no colour part. The chart part carries no
// format of its own, so the fonts, lines and fills of the chart come from here.
// Built from the chartStyle schema of [MS-ODRAWXML] (2.24.3: every element of the chart has an entry with a line, fill,
// effect and font reference). The values are a plain default of this library: dark gray text, light gray lines, the accent
// colours of the theme for the data. They are not a copy of one of the styles of Excel.
// (c) 2026 OraScope Co.,Ltd. GNU Lesser General Public License v3.0, as the rest of the library.
using System.Text;

namespace OfficeOpenXml.Drawing.Chart
{
    internal static class ExcelChartExStyle
    {
        const string NsStyle = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";

        // text: 65 % of the text colour; lines: 15 % of it
        const string TextColor = "<a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\" /><a:lumOff val=\"35000\" /></a:schemeClr>";
        const string LineColor = "<a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\" /><a:lumOff val=\"85000\" /></a:schemeClr>";
        const string ThinLine = "<a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill>" + LineColor + "</a:solidFill><a:round /></a:ln>";
        // the colour of the series or point, taken from the colour part
        const string DataColor = "<cs:styleClr val=\"auto\" />";
        const string DataFill = "<a:solidFill><a:schemeClr val=\"phClr\" /></a:solidFill>";

        /// <summary>
        /// The style part: one entry for every element a chart can have, in the order of the schema.
        /// </summary>
        /// <param name="lightLabels">The data labels are written on the fill of the point (treemap): light text</param>
        internal static string StyleXml(bool lightLabels)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<cs:chartStyle xmlns:cs=\"" + NsStyle + "\" xmlns:a=\"" + ExcelPackage.schemaDrawings + "\" id=\"1\">");
            Entry(sb, "axisTitle", size: 900);
            Entry(sb, "categoryAxis", shape: ThinLine, size: 900);
            Entry(sb, "chartArea", mods: "allowNoFillOverride allowNoLineOverride", fontColor: "<a:schemeClr val=\"tx1\" />",
                shape: "<a:solidFill><a:schemeClr val=\"bg1\" /></a:solidFill>" + ThinLine, size: 1000);
            if (lightLabels) Entry(sb, "dataLabel", fontColor: "<a:schemeClr val=\"lt1\" />", size: 900);
            else Entry(sb, "dataLabel", size: 900);
            Entry(sb, "dataLabelCallout", fontColor: "<a:schemeClr val=\"dk1\"><a:lumMod val=\"65000\" /><a:lumOff val=\"35000\" /></a:schemeClr>",
                shape: "<a:solidFill><a:schemeClr val=\"lt1\" /></a:solidFill><a:ln><a:solidFill><a:schemeClr val=\"dk1\"><a:lumMod val=\"25000\" /><a:lumOff val=\"75000\" /></a:schemeClr></a:solidFill></a:ln>", size: 900);
            Entry(sb, "dataPoint", fillColor: DataColor, shape: DataFill);
            Entry(sb, "dataPoint3D", fillColor: DataColor, shape: DataFill);
            Entry(sb, "dataPointLine", lineColor: DataColor, shape: "<a:ln w=\"28575\" cap=\"rnd\"><a:solidFill><a:schemeClr val=\"phClr\" /></a:solidFill><a:round /></a:ln>");
            Entry(sb, "dataPointMarker", lineColor: DataColor, fillColor: DataColor, shape: DataFill + "<a:ln w=\"9525\"><a:solidFill><a:schemeClr val=\"lt1\" /></a:solidFill></a:ln>");
            sb.Append("<cs:dataPointMarkerLayout symbol=\"circle\" size=\"5\" />");
            Entry(sb, "dataPointWireframe", lineColor: DataColor, shape: "<a:ln w=\"28575\" cap=\"rnd\"><a:solidFill><a:schemeClr val=\"phClr\" /></a:solidFill><a:round /></a:ln>");
            Entry(sb, "dataTable", shape: ThinLine, size: 900);
            Entry(sb, "downBar", fontColor: "<a:schemeClr val=\"dk1\" />", shape: "<a:solidFill><a:schemeClr val=\"dk1\"><a:lumMod val=\"65000\" /><a:lumOff val=\"35000\" /></a:schemeClr></a:solidFill>");
            Entry(sb, "dropLine", shape: ThinLine);
            Entry(sb, "errorBar", shape: ThinLine);
            Entry(sb, "floor");
            Entry(sb, "gridlineMajor", shape: ThinLine);
            Entry(sb, "gridlineMinor", shape: ThinLine);
            Entry(sb, "hiLoLine", shape: ThinLine);
            Entry(sb, "leaderLine", shape: ThinLine);
            Entry(sb, "legend", size: 900);
            Entry(sb, "plotArea", mods: "allowNoFillOverride allowNoLineOverride");
            Entry(sb, "plotArea3D", mods: "allowNoFillOverride allowNoLineOverride");
            Entry(sb, "seriesAxis", shape: ThinLine, size: 900);
            Entry(sb, "seriesLine", shape: ThinLine);
            Entry(sb, "title", size: 1400);
            Entry(sb, "trendline", lineColor: DataColor, shape: "<a:ln w=\"19050\" cap=\"rnd\"><a:solidFill><a:schemeClr val=\"phClr\" /></a:solidFill><a:prstDash val=\"sysDash\" /></a:ln>");
            Entry(sb, "trendlineLabel", size: 900);
            Entry(sb, "upBar", fontColor: "<a:schemeClr val=\"dk1\" />", shape: "<a:solidFill><a:schemeClr val=\"lt1\" /></a:solidFill>");
            Entry(sb, "valueAxis", size: 900);
            Entry(sb, "wall");
            sb.Append("</cs:chartStyle>");
            return sb.ToString();
        }

        // One entry: the references every entry must have, then the shape and the text size when it has them
        static void Entry(StringBuilder sb, string name, string mods = null, string lineColor = null, string fillColor = null, string fontColor = TextColor, string shape = null, int size = 0)
        {
            sb.Append("<cs:" + name + (mods == null ? "" : " mods=\"" + mods + "\"") + ">");
            sb.Append(lineColor == null ? "<cs:lnRef idx=\"0\" />" : "<cs:lnRef idx=\"0\">" + lineColor + "</cs:lnRef>");
            sb.Append(fillColor == null ? "<cs:fillRef idx=\"0\" />" : "<cs:fillRef idx=\"0\">" + fillColor + "</cs:fillRef>");
            sb.Append("<cs:effectRef idx=\"0\" />");
            sb.Append("<cs:fontRef idx=\"minor\">" + fontColor + "</cs:fontRef>");
            if (shape != null) sb.Append("<cs:spPr>" + shape + "</cs:spPr>");
            if (size > 0) sb.Append("<cs:defRPr sz=\"" + size.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\" />");
            sb.Append("</cs:" + name + ">");
        }

        /// <summary>
        /// The colour part: the six accent colours of the theme, one after the other, then again darker and lighter.
        /// </summary>
        internal static string ColorsXml()
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<cs:colorStyle xmlns:cs=\"" + NsStyle + "\" xmlns:a=\"" + ExcelPackage.schemaDrawings + "\" meth=\"cycle\" id=\"10\">");
            for (int i = 1; i <= 6; i++) sb.Append("<a:schemeClr val=\"accent" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\" />");
            sb.Append("<cs:variation />");
            sb.Append("<cs:variation><a:lumMod val=\"60000\" /></cs:variation>");
            sb.Append("<cs:variation><a:lumMod val=\"80000\" /><a:lumOff val=\"20000\" /></cs:variation>");
            sb.Append("</cs:colorStyle>");
            return sb.ToString();
        }
    }
}
