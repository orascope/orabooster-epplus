// orabooster-epplus: tests for the issues fixed in this fork (https://github.com/orascope/orabooster-epplus/issues)
using System;
using System.IO;
using System.Linq;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace EPPlusTest
{
    [TestClass]
    public class OraBoosterIssuesTests
    {
        static ExcelPackage Reload(ExcelPackage a_package)
        {
            var ms = new MemoryStream();
            a_package.SaveAs(ms);
            ms.Position = 0;
            return new ExcelPackage(ms);
        }

        static int[] TableIds(ExcelPackage a_package)
        {
            return a_package.Workbook.Worksheets.SelectMany(ws => ws.Tables).Select(t => t.Id).ToArray();
        }

        // #2: deleting a table on a workbook with several sheets left the next id at 0 or at an id in use
        [TestMethod]
        public void Table_ids_stay_unique_after_a_delete_with_several_sheets()
        {
            using (var p = new ExcelPackage())
            {
                for (int s = 1; s <= 3; s++)
                {
                    var ws = p.Workbook.Worksheets.Add("S" + s);
                    ws.Cells["A1"].Value = "H"; ws.Cells["A2"].Value = 1;
                    ws.Tables.Add(ws.Cells["A1:A2"], "Table_S" + s);
                }
                p.Workbook.Worksheets["S1"].Tables.Delete("Table_S1");
                var ws1 = p.Workbook.Worksheets["S1"];
                ws1.Cells["C1"].Value = "H"; ws1.Cells["C2"].Value = 2;
                var added = ws1.Tables.Add(ws1.Cells["C1:C2"], "Table_New");

                int[] ids = TableIds(p);
                CollectionAssert.AllItemsAreUnique(ids, string.Join(",", ids));
                Assert.IsTrue(ids.All(i => i > 0), string.Join(",", ids));

                using (var r = Reload(p))
                {
                    int[] reloaded = TableIds(r);
                    Assert.AreEqual(3, reloaded.Length);
                    CollectionAssert.AllItemsAreUnique(reloaded, string.Join(",", reloaded));
                    Assert.IsTrue(reloaded.All(i => i > 0), string.Join(",", reloaded));
                    Assert.IsNotNull(r.Workbook.Worksheets["S1"].Tables["Table_New"]);
                }
            }
        }

        [TestMethod]
        public void Deleting_every_table_and_adding_again_starts_at_one()
        {
            using (var p = new ExcelPackage())
            {
                var a = p.Workbook.Worksheets.Add("A");
                var b = p.Workbook.Worksheets.Add("B");
                a.Cells["A1"].Value = "H"; b.Cells["A1"].Value = "H";
                a.Tables.Add(a.Cells["A1:A2"], "Table_A");
                b.Tables.Add(b.Cells["A1:A2"], "Table_B");
                a.Tables.Delete("Table_A");
                b.Tables.Delete("Table_B");
                Assert.AreEqual(1, a.Tables.Add(a.Cells["A1:A2"], "Table_A2").Id);
                Assert.AreEqual(2, b.Tables.Add(b.Cells["A1:A2"], "Table_B2").Id);
            }
        }

        static ExcelChart AddChartWithSeries(ExcelWorksheet a_ws, eChartType a_type, int a_series)
        {
            a_ws.Cells["A1"].Value = "Day";
            string[] names = { "Open", "High", "Low", "Close" };
            for (int c = 0; c < a_series; c++) a_ws.Cells[1, c + 2].Value = names[names.Length - a_series + c];
            for (int r = 2; r <= 6; r++)
            {
                a_ws.Cells[r, 1].Value = "d" + r;
                for (int c = 0; c < a_series; c++) a_ws.Cells[r, c + 2].Value = 10 + r + c;
            }
            var chart = a_ws.Drawings.AddChart("C1", a_type);
            for (int c = 0; c < a_series; c++)
            {
                var s = chart.Series.Add(ExcelCellBase.GetAddress(2, c + 2, 6, c + 2), "A2:A6");
                s.HeaderAddress = a_ws.Cells[1, c + 2];
            }
            return chart;
        }

        static XmlNode PlotChild(ExcelChart a_chart, string a_localName)
        {
            var ns = new XmlNamespaceManager(a_chart.ChartXml.NameTable);
            ns.AddNamespace("c", ExcelPackage.schemaChart);
            return a_chart.ChartXml.SelectSingleNode("//c:plotArea/c:" + a_localName, ns);
        }

        // #3: AddChart threw NotImplementedException for Column3D
        [TestMethod]
        public void Column3D_chart_is_a_3d_bar_chart()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("S");
                var chart = AddChartWithSeries(ws, eChartType.Column3D, 2);
                var bar = PlotChild(chart, "bar3DChart");
                Assert.IsNotNull(bar);
                using (var r = Reload(p))
                {
                    Assert.AreEqual(1, r.Workbook.Worksheets["S"].Drawings.Count);
                }
            }
        }

        // #3: stock charts (high-low-close, open-high-low-close)
        [TestMethod]
        public void Stock_charts_have_high_low_lines_and_no_vary_colors()
        {
            foreach (var type in new[] { eChartType.StockHLC, eChartType.StockOHLC })
            {
                using (var p = new ExcelPackage())
                {
                    var ws = p.Workbook.Worksheets.Add("S");
                    int series = type == eChartType.StockOHLC ? 4 : 3;
                    var chart = AddChartWithSeries(ws, type, series);
                    var stock = PlotChild(chart, "stockChart");
                    Assert.IsNotNull(stock, type.ToString());
                    var children = stock.ChildNodes.Cast<XmlNode>().Select(n => n.LocalName).ToArray();
                    // CT_StockChart: ser{3,4}, dLbls?, dropLines?, hiLowLines?, upDownBars?, axId{2}
                    string order = string.Join(",", children);
                    Assert.AreEqual(series, children.Count(n => n == "ser"), order);
                    Assert.IsFalse(children.Contains("varyColors"), order);
                    Assert.IsTrue(Array.IndexOf(children, "hiLowLines") > Array.LastIndexOf(children, "ser"), order);
                    Assert.IsTrue(Array.IndexOf(children, "axId") > Array.IndexOf(children, "hiLowLines"), order);
                    Assert.AreEqual(type == eChartType.StockOHLC, children.Contains("upDownBars"), order);
                    using (var r = Reload(p))
                    {
                        Assert.AreEqual(1, r.Workbook.Worksheets["S"].Drawings.Count, type.ToString());
                        // Excel turns on the high-low lines only when the series carry cached values when the file is opened
                        var saved = (ExcelChart)r.Workbook.Worksheets["S"].Drawings[0];
                        var ns = new XmlNamespaceManager(saved.ChartXml.NameTable);
                        ns.AddNamespace("c", ExcelPackage.schemaChart);
                        var caches = saved.ChartXml.SelectNodes("//c:stockChart/c:ser/c:val/c:numRef/c:numCache", ns);
                        Assert.AreEqual(series, caches.Count, type.ToString());
                        foreach (XmlNode cache in caches) Assert.AreEqual(5, cache.SelectNodes("c:pt", ns).Count, type.ToString());
                    }
                }
            }
        }

        // #3: the volume stock charts are not supported; the check must come before anything is added to the drawing
        [TestMethod]
        public void Volume_stock_charts_throw_before_adding_a_drawing()
        {
            foreach (var type in new[] { eChartType.StockVHLC, eChartType.StockVOHLC })
            {
                using (var p = new ExcelPackage())
                {
                    var ws = p.Workbook.Worksheets.Add("S");
                    Assert.ThrowsException<NotSupportedException>(() => ws.Drawings.AddChart("C1", type), type.ToString());
                    Assert.AreEqual(0, ws.Drawings.Count, type.ToString());
                    ws.Cells["A1"].Value = 1;
                    using (var r = Reload(p))
                    {
                        Assert.AreEqual(0, r.Workbook.Worksheets["S"].Drawings.Count, type.ToString());
                    }
                }
            }
        }
    }
}
