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

        // A sheet with a day column, a volume column and the price columns (open,) high, low, close in rows 2 to 6
        static ExcelWorksheet StockSheet(ExcelPackage a_package, bool a_open)
        {
            var ws = a_package.Workbook.Worksheets.Add("S");
            string[] heads = a_open ? new[] { "Day", "Volume", "Open", "High", "Low", "Close" } : new[] { "Day", "Volume", "High", "Low", "Close" };
            for (int c = 0; c < heads.Length; c++) ws.Cells[1, c + 1].Value = heads[c];
            for (int r = 2; r <= 6; r++)
            {
                ws.Cells[r, 1].Value = "D" + r;
                ws.Cells[r, 2].Value = 1000 * r;
                double[] v = a_open ? new double[] { 100 + r, 110 + r, 90 + r, 105 + r } : new double[] { 110 + r, 90 + r, 105 + r };
                for (int c = 0; c < v.Length; c++) ws.Cells[r, c + 3].Value = v[c];
            }
            return ws;
        }

        // #11: a volume stock chart is a column chart (the volume, primary axes) and a stock chart (secondary axes).
        // Series added one after the other to the chart from AddChart go to the right part.
        [TestMethod]
        public void Volume_stock_charts_are_a_column_chart_and_a_stock_chart_on_the_secondary_axes()
        {
            foreach (var type in new[] { eChartType.StockVHLC, eChartType.StockVOHLC })
            {
                bool open = type == eChartType.StockVOHLC;
                int prices = open ? 4 : 3;
                eChartType stockType = open ? eChartType.StockOHLC : eChartType.StockHLC;
                using (var p = new ExcelPackage())
                {
                    var ws = StockSheet(p, open);
                    var chart = ws.Drawings.AddChart("C1", type);
                    for (int c = 2; c <= prices + 2; c++)
                    {
                        var s = chart.Series.Add(ExcelCellBase.GetAddress(2, c, 6, c), "A2:A6");
                        s.HeaderAddress = ws.Cells[1, c];
                    }
                    Assert.AreEqual(1, ws.Drawings.Count, type.ToString());
                    Assert.AreEqual(eChartType.ColumnClustered, chart.ChartType, type.ToString());
                    Assert.AreEqual(1, chart.Series.Count, type.ToString());
                    Assert.AreEqual(2, chart.PlotArea.ChartTypes.Count, type.ToString());
                    ExcelChart stock = chart.PlotArea.ChartTypes[1];
                    Assert.AreEqual(stockType, stock.ChartType, type.ToString());
                    Assert.AreEqual(prices, stock.Series.Count, type.ToString());
                    Assert.IsTrue(stock.UseSecondaryAxis, type.ToString());

                    XmlElement space = SavedSheet(p, "xl/charts/chart1.xml");
                    var ns = new XmlNamespaceManager(space.OwnerDocument.NameTable);
                    ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
                    XmlNode plot = space.SelectSingleNode("c:chart/c:plotArea", ns);
                    string order = string.Join(",", plot.ChildNodes.OfType<XmlElement>().Select(e => e.LocalName));
                    Assert.AreEqual("layout,barChart,stockChart,catAx,valAx,catAx,valAx", order, type.ToString());

                    // the volume on the first pair of axes, the prices on the second
                    string[] barAx = plot.SelectNodes("c:barChart/c:axId/@val", ns).OfType<XmlNode>().Select(n => n.Value).ToArray();
                    string[] stockAx = plot.SelectNodes("c:stockChart/c:axId/@val", ns).OfType<XmlNode>().Select(n => n.Value).ToArray();
                    string[] axes = plot.SelectNodes("c:catAx/c:axId/@val|c:valAx/c:axId/@val", ns).OfType<XmlNode>().Select(n => n.Value).ToArray();
                    CollectionAssert.AreEqual(new[] { axes[0], axes[1] }, barAx, type.ToString());
                    CollectionAssert.AreEquivalent(new[] { axes[2], axes[3] }, stockAx, type.ToString());
                    Assert.AreEqual(4, axes.Distinct().Count(), type.ToString());
                    XmlNode secondValAx = plot.SelectNodes("c:valAx", ns)[1];
                    Assert.AreEqual("max", secondValAx.SelectSingleNode("c:crosses/@val", ns).Value, type.ToString());
                    Assert.AreEqual("r", secondValAx.SelectSingleNode("c:axPos/@val", ns).Value, type.ToString());
                    Assert.AreEqual("1", plot.SelectNodes("c:catAx", ns)[1].SelectSingleNode("c:delete/@val", ns).Value, type.ToString());

                    // series: the volume is 0 on the column chart, the prices follow on the stock chart
                    Assert.AreEqual("0", string.Join(",", plot.SelectNodes("c:barChart/c:ser/c:idx/@val", ns).OfType<XmlNode>().Select(n => n.Value)), type.ToString());
                    Assert.AreEqual(open ? "1,2,3,4" : "1,2,3", string.Join(",", plot.SelectNodes("c:stockChart/c:ser/c:idx/@val", ns).OfType<XmlNode>().Select(n => n.Value)), type.ToString());
                    Assert.AreEqual(open ? "1,2,3,4" : "1,2,3", string.Join(",", plot.SelectNodes("c:stockChart/c:ser/c:order/@val", ns).OfType<XmlNode>().Select(n => n.Value)), type.ToString());
                    Assert.AreEqual("'S'!B2:B6", plot.SelectSingleNode("c:barChart/c:ser/c:val/c:numRef/c:f", ns).InnerText, type.ToString());
                    Assert.AreEqual("'S'!C2:C6", plot.SelectSingleNode("c:stockChart/c:ser/c:val/c:numRef/c:f", ns).InnerText, type.ToString());

                    Assert.IsNotNull(plot.SelectSingleNode("c:stockChart/c:hiLowLines", ns), type.ToString());
                    Assert.AreEqual(open, plot.SelectSingleNode("c:stockChart/c:upDownBars", ns) != null, type.ToString());
                    Assert.IsNull(plot.SelectSingleNode("c:stockChart/c:varyColors", ns), type.ToString());
                    // the close of a high-low-close chart is a dash: the third series of the stock chart, whatever its idx
                    string symbols = string.Join(",", plot.SelectNodes("c:stockChart/c:ser/c:marker/c:symbol/@val", ns).OfType<XmlNode>().Select(n => n.Value));
                    Assert.AreEqual(open ? "none,none,none,none" : "none,none,dash", symbols, type.ToString());
                    // Excel turns the high-low lines on only when the stock series carry cached values
                    foreach (XmlNode cache in plot.SelectNodes("c:stockChart/c:ser/c:val/c:numRef/c:numCache", ns)) Assert.AreEqual(5, cache.SelectNodes("c:pt", ns).Count, type.ToString());
                    Assert.AreEqual(prices, plot.SelectNodes("c:stockChart/c:ser/c:val/c:numRef/c:numCache", ns).Count, type.ToString());

                    using (var r = Reload(p))
                    {
                        var rc = (ExcelChart)r.Workbook.Worksheets["S"].Drawings[0];
                        ExcelChart[] types = rc.PlotArea.ChartTypes.ToArray(); // each chart type once (issue #12)
                        Assert.AreEqual(2, types.Length, type.ToString());
                        Assert.AreEqual(1, rc.Series.Count, type.ToString());
                        Assert.AreEqual(stockType, types[1].ChartType, type.ToString());
                        Assert.AreEqual(prices, types[1].Series.Count, type.ToString());

                        // saved again after a load: still one column chart and one stock chart with the same series
                        XmlNode plot2 = SavedSheet(r, "xl/charts/chart1.xml").SelectSingleNode("c:chart/c:plotArea", ns);
                        Assert.AreEqual(order, string.Join(",", plot2.ChildNodes.OfType<XmlElement>().Select(e => e.LocalName)), type.ToString());
                        Assert.AreEqual(prices, plot2.SelectNodes("c:stockChart/c:ser", ns).Count, type.ToString());
                        Assert.AreEqual(prices, plot2.SelectNodes("c:stockChart/c:ser/c:val/c:numRef/c:numCache", ns).Count, type.ToString());
                    }
                }
            }
        }

        // #11: a pivot chart can not be combined with another chart type; that is refused before anything is added to the drawing
        [TestMethod]
        public void Volume_stock_chart_as_a_pivot_chart_is_refused_before_adding_a_drawing()
        {
            using (var p = new ExcelPackage())
            {
                var ws = StockSheet(p, false);
                var pivot = ws.PivotTables.Add(ws.Cells["H1"], ws.Cells["A1:E6"], "Pivot1");
                Assert.ThrowsException<NotSupportedException>(() => ws.Drawings.AddChart("C1", eChartType.StockVHLC, pivot));
                Assert.AreEqual(0, ws.Drawings.Count);
                using (var r = Reload(p))
                {
                    Assert.AreEqual(0, r.Workbook.Worksheets["S"].Drawings.Count);
                }
            }
        }

        // #11: the same on a chart sheet, with the series given as ranges of another sheet
        [TestMethod]
        public void Volume_stock_chart_on_a_chart_sheet()
        {
            using (var p = new ExcelPackage())
            {
                var ws = StockSheet(p, true);
                ExcelChart chart = p.Workbook.Worksheets.AddChart("CS", eChartType.StockVOHLC).Chart;
                for (int c = 2; c <= 6; c++) chart.Series.Add(ws.Cells[2, c, 6, c], ws.Cells["A2:A6"]);
                Assert.AreEqual(1, chart.Series.Count);
                Assert.AreEqual(4, chart.PlotArea.ChartTypes[1].Series.Count);

                XmlElement space = SavedSheet(p, "xl/charts/chart1.xml");
                var ns = new XmlNamespaceManager(space.OwnerDocument.NameTable);
                ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
                Assert.AreEqual(1, space.SelectNodes("c:chart/c:plotArea/c:barChart/c:ser", ns).Count);
                Assert.AreEqual(4, space.SelectNodes("c:chart/c:plotArea/c:stockChart/c:ser", ns).Count);
                // the caches come from the sheet the series refer to
                Assert.AreEqual("102", space.SelectSingleNode("c:chart/c:plotArea/c:stockChart/c:ser/c:val/c:numRef/c:numCache/c:pt/c:v", ns).InnerText);
            }
        }

        // #11: a copy of the worksheet has the whole chart, both parts
        [TestMethod]
        public void Volume_stock_chart_is_copied_with_the_worksheet()
        {
            using (var p = new ExcelPackage())
            {
                var ws = StockSheet(p, false);
                var chart = ws.Drawings.AddChart("C1", eChartType.StockVHLC);
                for (int c = 2; c <= 5; c++) chart.Series.Add(ExcelCellBase.GetAddress(2, c, 6, c), "A2:A6");
                var copy = p.Workbook.Worksheets.Add("Copy", ws);
                Assert.AreEqual(1, copy.Drawings.Count);
                using (var r = Reload(p))
                {
                    foreach (string name in new[] { "S", "Copy" })
                    {
                        ExcelChart[] types = ((ExcelChart)r.Workbook.Worksheets[name].Drawings[0]).PlotArea.ChartTypes.ToArray();
                        Assert.AreEqual(2, types.Length, name);
                        Assert.AreEqual(eChartType.ColumnClustered, types[0].ChartType, name);
                        Assert.AreEqual(1, types[0].Series.Count, name);
                        Assert.AreEqual(eChartType.StockHLC, types[1].ChartType, name);
                        Assert.AreEqual(3, types[1].Series.Count, name);
                    }
                }
            }
        }

        // #11: a loaded stock chart with up/down bars is an open-high-low-close chart (it was reported as high-low-close)
        [TestMethod]
        public void Loaded_stock_chart_with_an_open_series_is_reported_as_StockOHLC()
        {
            foreach (var type in new[] { eChartType.StockHLC, eChartType.StockOHLC })
            {
                using (var p = new ExcelPackage())
                {
                    var ws = StockSheet(p, type == eChartType.StockOHLC);
                    var chart = ws.Drawings.AddChart("C1", type);
                    for (int c = 3; c <= (type == eChartType.StockOHLC ? 6 : 5); c++) chart.Series.Add(ExcelCellBase.GetAddress(2, c, 6, c), "A2:A6");
                    using (var r = Reload(p))
                    {
                        Assert.AreEqual(type, ((ExcelChart)r.Workbook.Worksheets["S"].Drawings[0]).ChartType);
                    }
                }
            }
        }

        // The worksheet part as it is saved in the package
        static XmlElement SavedSheet(ExcelPackage a_package, string a_entry = "xl/worksheets/sheet1.xml")
        {
            var ms = new MemoryStream();
            a_package.SaveAs(ms);
            ms.Position = 0;
            using (var zip = new System.IO.Compression.ZipArchive(ms))
            using (var s = zip.GetEntry(a_entry).Open())
            {
                var doc = new XmlDocument();
                doc.Load(s);
                return doc.DocumentElement;
            }
        }

        static string[] ChildNames(XmlElement a_sheet)
        {
            return a_sheet.ChildNodes.OfType<XmlElement>().Select(e => e.LocalName).ToArray();
        }

        // #9: ignored errors are written at their place in the schema order and read back
        [TestMethod]
        public void Ignored_errors_are_saved_in_schema_order_and_read_back()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("S");
                ws.Cells["A1"].Value = "Id"; ws.Cells["B1"].Value = "N";
                for (int r = 2; r <= 5; r++) { ws.Cells[r, 1].Value = "00" + r; ws.Cells[r, 2].Value = r; }
                // parts that come after ignoredErrors in the schema, made before and after it
                ws.Tables.Add(ws.Cells["A1:B5"], "Table_S");
                ws.IgnoredErrors.Add("A2:A5", eErrorChecks.NumberStoredAsText);
                ws.IgnoredErrors.Add("D:D,F2:G3", eErrorChecks.Formula | eErrorChecks.EvalError | eErrorChecks.CalculatedColumn);
                var chart = ws.Drawings.AddChart("C1", eChartType.Line);
                chart.Series.Add("B2:B5", "A2:A5");
                ws.PrinterSettings.TopMargin = 1m;
                ws.Cells["H1"].AddComment("c", "a");

                XmlElement sheet = SavedSheet(p);
                string[] names = ChildNames(sheet);
                int ix = Array.IndexOf(names, "ignoredErrors");
                Assert.IsTrue(ix > 0, string.Join(",", names));
                Assert.AreEqual(1, names.Count(n => n == "ignoredErrors"));
                foreach (string before in new[] { "sheetData", "pageMargins" })
                    Assert.IsTrue(Array.IndexOf(names, before) >= 0 && Array.IndexOf(names, before) < ix, before + ": " + string.Join(",", names));
                foreach (string after in new[] { "drawing", "legacyDrawing", "tableParts" })
                    Assert.IsTrue(Array.IndexOf(names, after) > ix, after + ": " + string.Join(",", names));

                var items = sheet.ChildNodes.OfType<XmlElement>().Single(e => e.LocalName == "ignoredErrors").ChildNodes.OfType<XmlElement>().ToArray();
                Assert.AreEqual(2, items.Length);
                Assert.AreEqual("<ignoredError sqref=\"A2:A5\" numberStoredAsText=\"1\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" />", items[0].OuterXml);
                Assert.AreEqual("<ignoredError sqref=\"D1:D1048576 F2:G3\" evalError=\"1\" formula=\"1\" calculatedColumn=\"1\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" />", items[1].OuterXml);

                using (var r = Reload(p))
                {
                    var rs = r.Workbook.Worksheets["S"];
                    Assert.AreEqual(2, rs.IgnoredErrors.Count);
                    Assert.AreEqual("A2:A5", rs.IgnoredErrors[0].Address);
                    Assert.AreEqual(eErrorChecks.NumberStoredAsText, rs.IgnoredErrors[0].Checks);
                    Assert.AreEqual("D1:D1048576 F2:G3", rs.IgnoredErrors[1].Address);
                    Assert.AreEqual(eErrorChecks.Formula | eErrorChecks.EvalError | eErrorChecks.CalculatedColumn, rs.IgnoredErrors[1].Checks);

                    // changed after a load, saved again
                    rs.IgnoredErrors[0].Checks = eErrorChecks.All;
                    rs.IgnoredErrors[1].Address = "C3";
                    rs.IgnoredErrors[1].Checks = eErrorChecks.TwoDigitTextYear;
                    using (var r2 = Reload(r))
                    {
                        var rs2 = r2.Workbook.Worksheets["S"];
                        Assert.AreEqual(eErrorChecks.All, rs2.IgnoredErrors[0].Checks);
                        Assert.AreEqual("C3", rs2.IgnoredErrors[1].Address);
                        Assert.AreEqual(eErrorChecks.TwoDigitTextYear, rs2.IgnoredErrors[1].Checks);
                        Assert.AreEqual(1, r2.Workbook.Worksheets["S"].Tables.Count);
                        Assert.AreEqual(1, r2.Workbook.Worksheets["S"].Drawings.Count);
                    }
                }
            }
        }

        // #9: a workbook made by Excel writes "true"/"false" or "1"/"0"; entries that are not touched are kept as they are
        [TestMethod]
        public void Ignored_errors_of_a_loaded_workbook_are_read_and_kept()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("S");
                ws.Cells["A1"].Value = 1;
                var e = ws.IgnoredErrors.Add("A1", eErrorChecks.Formula);
                // as another writer could have saved it
                e.Node.SetAttribute("formula", "true");
                e.Node.SetAttribute("evalError", "0");
                e.Node.SetAttribute("numberStoredAsText", "false");
                e.Node.SetAttribute("unlockedFormula", "1");
                using (var r = Reload(p))
                {
                    var item = r.Workbook.Worksheets["S"].IgnoredErrors[0];
                    Assert.AreEqual(eErrorChecks.Formula | eErrorChecks.UnlockedFormula, item.Checks);
                    XmlElement saved = SavedSheet(r).ChildNodes.OfType<XmlElement>().Single(n => n.LocalName == "ignoredErrors");
                    Assert.AreEqual("true", ((XmlElement)saved.FirstChild).GetAttribute("formula"));
                    Assert.AreEqual("false", ((XmlElement)saved.FirstChild).GetAttribute("numberStoredAsText"));
                }
            }
        }

        // #9: a sheet without ignored errors has no element, also after the collection was read, emptied, or rows were moved
        [TestMethod]
        public void Ignored_errors_element_is_absent_when_there_are_none()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("S");
                ws.Cells["A1"].Value = 1;
                Assert.AreEqual(0, ws.IgnoredErrors.Count);
                ws.InsertRow(1, 1);
                ws.DeleteColumn(3, 1);
                CollectionAssert.DoesNotContain(ChildNames(SavedSheet(p)), "ignoredErrors");

                var a = ws.IgnoredErrors.Add("A1:A3", eErrorChecks.NumberStoredAsText);
                var b = ws.IgnoredErrors.Add(ws.Cells["B1:B3"], eErrorChecks.EvalError);
                Assert.AreEqual("B1:B3", b.Address);
                Assert.IsTrue(ws.IgnoredErrors.Remove(a));
                Assert.IsFalse(ws.IgnoredErrors.Remove(a));
                Assert.AreEqual(1, ws.IgnoredErrors.Count);
                CollectionAssert.Contains(ChildNames(SavedSheet(p)), "ignoredErrors");
                Assert.IsTrue(ws.IgnoredErrors.Remove(b));
                CollectionAssert.DoesNotContain(ChildNames(SavedSheet(p)), "ignoredErrors");

                ws.IgnoredErrors.Add("A1", eErrorChecks.All);
                ws.IgnoredErrors.Add("A2", eErrorChecks.All);
                ws.IgnoredErrors.Clear();
                Assert.AreEqual(0, ws.IgnoredErrors.Count);
                CollectionAssert.DoesNotContain(ChildNames(SavedSheet(p)), "ignoredErrors");

                // usable again after it was emptied
                ws.IgnoredErrors.Add("A1", eErrorChecks.All);
                Assert.AreEqual(1, SavedSheet(p).ChildNodes.OfType<XmlElement>().Count(n => n.LocalName == "ignoredErrors"));
            }
        }

        // #9: inserting and deleting rows and columns moves the ranges; a range whose cells are all deleted goes away
        [TestMethod]
        public void Ignored_errors_move_with_inserted_and_deleted_rows_and_columns()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("S");
                for (int r = 1; r <= 20; r++) for (int c = 1; c <= 8; c++) ws.Cells[r, c].Value = r * c;
                var a = ws.IgnoredErrors.Add("B5:C10 E5:E10", eErrorChecks.NumberStoredAsText);
                var b = ws.IgnoredErrors.Add("G12:G13", eErrorChecks.Formula);

                ws.InsertRow(1, 2);          // above: both move down
                Assert.AreEqual("B7:C12 E7:E12", a.Address);
                Assert.AreEqual("G14:G15", b.Address);
                ws.InsertRow(8, 1);          // inside a: it grows, b moves
                Assert.AreEqual("B7:C13 E7:E13", a.Address);
                Assert.AreEqual("G15:G16", b.Address);
                ws.InsertRow(30, 5);         // below: nothing moves
                Assert.AreEqual("B7:C13 E7:E13", a.Address);

                ws.InsertColumn(1, 1);       // left of both
                Assert.AreEqual("C7:D13 F7:F13", a.Address);
                Assert.AreEqual("H15:H16", b.Address);
                ws.InsertColumn(4, 2);       // inside the first range of a
                Assert.AreEqual("C7:F13 H7:H13", a.Address);
                Assert.AreEqual("J15:J16", b.Address);

                ws.DeleteRow(1, 2);          // above
                Assert.AreEqual("C5:F11 H5:H11", a.Address);
                Assert.AreEqual("J13:J14", b.Address);
                ws.DeleteRow(6, 2);          // inside a
                Assert.AreEqual("C5:F9 H5:H9", a.Address);
                Assert.AreEqual("J11:J12", b.Address);

                ws.DeleteColumn(8, 1);       // the second range of a goes away, the first stays
                Assert.AreEqual("C5:F9", a.Address);
                Assert.AreEqual("I11:I12", b.Address);
                ws.DeleteColumn(1, 1);       // left
                Assert.AreEqual("B5:E9", a.Address);
                Assert.AreEqual("H11:H12", b.Address);
                ws.DeleteColumn(3, 1);       // inside a
                Assert.AreEqual("B5:D9", a.Address);

                ws.DeleteRow(11, 2);         // every cell of b
                Assert.AreEqual(1, ws.IgnoredErrors.Count);
                Assert.AreSame(a, ws.IgnoredErrors[0]);
                ws.DeleteRow(4, 10);         // every cell of a: the element goes away
                Assert.AreEqual(0, ws.IgnoredErrors.Count);
                CollectionAssert.DoesNotContain(ChildNames(SavedSheet(p)), "ignoredErrors");
            }

            // the same on a workbook that was loaded and whose collection was not read first
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("S");
                ws.Cells["A1"].Value = 1; ws.Cells["A9"].Value = 2;
                ws.IgnoredErrors.Add("A5:A9", eErrorChecks.NumberStoredAsText);
                using (var r = Reload(p))
                {
                    var rs = r.Workbook.Worksheets["S"];
                    rs.InsertRow(2, 3);
                    Assert.AreEqual("A8:A12", rs.IgnoredErrors[0].Address);
                }
            }
        }

        // #9: wrong arguments are refused before anything is added
        [TestMethod]
        public void Ignored_errors_refuse_wrong_arguments()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("S");
                ws.Cells["A1"].Value = 1;
                Assert.ThrowsException<ArgumentException>(() => ws.IgnoredErrors.Add("", eErrorChecks.All));
                Assert.ThrowsException<ArgumentException>(() => ws.IgnoredErrors.Add("Other!A1", eErrorChecks.All));
                Assert.ThrowsException<ArgumentException>(() => ws.IgnoredErrors.Add("A1", eErrorChecks.None));
                Assert.ThrowsException<ArgumentNullException>(() => ws.IgnoredErrors.Add((ExcelAddressBase)null, eErrorChecks.All));
                Assert.ThrowsException<ArgumentException>(() => ws.IgnoredErrors.Add("NotARange", eErrorChecks.All));
                Assert.ThrowsException<ArgumentException>(() => ws.IgnoredErrors.Add("A1 NotARange", eErrorChecks.All));
                Assert.AreEqual(0, ws.IgnoredErrors.Count);
                CollectionAssert.DoesNotContain(ChildNames(SavedSheet(p)), "ignoredErrors");

                // a chart sheet has no cells
                var chartSheet = p.Workbook.Worksheets.AddChart("Chart", eChartType.Line);
                Assert.ThrowsException<NotSupportedException>(() => chartSheet.IgnoredErrors.Count);
            }
        }

        // #9: a copy of a worksheet has its own copy of the ignored errors
        [TestMethod]
        public void Ignored_errors_are_copied_with_the_worksheet()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("S");
                ws.Cells["A1"].Value = "001";
                ws.IgnoredErrors.Add("A1:A9", eErrorChecks.NumberStoredAsText);
                var copy = p.Workbook.Worksheets.Add("Copy", ws);
                Assert.AreEqual(1, copy.IgnoredErrors.Count);
                copy.IgnoredErrors[0].Address = "B1:B9";
                copy.IgnoredErrors.Add("C1", eErrorChecks.EvalError);
                Assert.AreEqual("A1:A9", ws.IgnoredErrors[0].Address);
                Assert.AreEqual(1, ws.IgnoredErrors.Count);
                using (var r = Reload(p))
                {
                    Assert.AreEqual("A1:A9", r.Workbook.Worksheets["S"].IgnoredErrors[0].Address);
                    Assert.AreEqual(2, r.Workbook.Worksheets["Copy"].IgnoredErrors.Count);
                    Assert.AreEqual("B1:B9", r.Workbook.Worksheets["Copy"].IgnoredErrors[0].Address);
                }
            }
        }

        // #12: a loaded chart has each of its chart types once in PlotArea.ChartTypes, and a series added to a loaded chart
        // gets the next number without giving the series that were there new ones
        [TestMethod]
        public void Loaded_chart_lists_each_chart_type_once()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("S");
                for (int r = 1; r <= 6; r++) for (int c = 1; c <= 6; c++) ws.Cells[r, c].Value = r == 1 ? (object)("H" + c) : r * c;
                var single = ws.Drawings.AddChart("Single", eChartType.ColumnClustered);
                single.Series.Add("B2:B6", "A2:A6");
                single.Series.Add("C2:C6", "A2:A6");
                var combined = ws.Drawings.AddChart("Combined", eChartType.ColumnClustered);
                combined.Series.Add("B2:B6", "A2:A6");
                combined.Series.Add("C2:C6", "A2:A6");
                var line = combined.PlotArea.ChartTypes.Add(eChartType.Line);
                line.Series.Add("D2:D6", "A2:A6");
                Assert.AreEqual(1, single.PlotArea.ChartTypes.Count);
                Assert.AreEqual(2, combined.PlotArea.ChartTypes.Count);

                using (var r = Reload(p))
                {
                    var rs = (ExcelChart)r.Workbook.Worksheets["S"].Drawings["Single"];
                    var rc = (ExcelChart)r.Workbook.Worksheets["S"].Drawings["Combined"];
                    Assert.AreEqual(1, rs.PlotArea.ChartTypes.Count, "one chart type");
                    Assert.AreSame(rs, rs.PlotArea.ChartTypes[0]);
                    Assert.AreEqual(2, rc.PlotArea.ChartTypes.Count, "two chart types");
                    Assert.AreSame(rc, rc.PlotArea.ChartTypes[0]);
                    Assert.AreEqual(eChartType.Line, rc.PlotArea.ChartTypes[1].ChartType);
                    Assert.AreEqual(1, rc.PlotArea.ChartTypes[1].Series.Count);

                    // one more series on each: the numbers go on, the ones that were there keep theirs
                    rs.Series.Add("D2:D6", "A2:A6");
                    rc.Series.Add("E2:E6", "A2:A6");
                    rc.PlotArea.ChartTypes[1].Series.Add("F2:F6", "A2:A6");
                    XmlElement one = SavedSheet(r, "xl/charts/chart1.xml");
                    var ns = new XmlNamespaceManager(one.OwnerDocument.NameTable);
                    ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
                    Assert.AreEqual("0,1,2", string.Join(",", one.SelectNodes("//c:ser/c:idx/@val", ns).OfType<XmlNode>().Select(n => n.Value)));
                    Assert.AreEqual("0,1,2", string.Join(",", one.SelectNodes("//c:ser/c:order/@val", ns).OfType<XmlNode>().Select(n => n.Value)));
                    Assert.AreEqual("'S'!B2:B6|'S'!C2:C6|'S'!D2:D6", string.Join("|", one.SelectNodes("//c:ser/c:val/c:numRef/c:f", ns).OfType<XmlNode>().Select(n => n.InnerText)));
                    XmlElement two = SavedSheet(r, "xl/charts/chart2.xml");
                    string[] ids = two.SelectNodes("//c:ser/c:idx/@val", ns).OfType<XmlNode>().Select(n => n.Value).ToArray();
                    Assert.AreEqual(5, ids.Length);
                    CollectionAssert.AreEquivalent(new[] { "0", "1", "2", "3", "4" }, ids, string.Join(",", ids));
                    Assert.AreEqual(3, two.SelectNodes("//c:barChart/c:ser", ns).Count);
                    Assert.AreEqual(2, two.SelectNodes("//c:lineChart/c:ser", ns).Count);
                }
            }
        }

        // #15: a number that is not a date in a cell with a date format shows as # signs, as in Excel. It made Text and
        // AutoFitColumns throw, and one such cell stopped the autofit of the whole range.
        [TestMethod]
        public void Number_that_is_not_a_date_in_a_date_format_does_not_throw()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("S");
                ws.Cells["A1"].Value = -700000;       // before the first date .NET has
                ws.Cells["A2"].Value = 3000000;       // after the last date
                ws.Cells["A3"].Value = double.NaN;
                ws.Cells["A4"].Value = 46000;         // a date
                ws.Cells["A5"].Value = -1.5;          // before 1900: Excel shows # signs, the library shows the date as before
                ws.Cells["B1"].Value = "a text that is longer than the dates";
                ws.Cells["A1:A5"].Style.Numberformat.Format = "yyyy-mm-dd";
                Assert.AreEqual("########", ws.Cells["A1"].Text);
                Assert.AreEqual("########", ws.Cells["A2"].Text);
                Assert.AreEqual("########", ws.Cells["A3"].Text);
                Assert.AreEqual("2025-12-09", ws.Cells["A4"].Text);
                Assert.AreEqual("1899-12-29", ws.Cells["A5"].Text);

                ws.Cells["A1:B5"].AutoFitColumns(0);
                Assert.IsTrue(ws.Column(1).Width > 8, ws.Column(1).Width.ToString());
                Assert.IsTrue(ws.Column(2).Width > ws.Column(1).Width, "the other column was fitted too");
                using (var r = Reload(p))
                {
                    Assert.AreEqual(-700000d, Convert.ToDouble(r.Workbook.Worksheets["S"].Cells["A1"].Value));
                    Assert.AreEqual("########", r.Workbook.Worksheets["S"].Cells["A2"].Text);
                }
            }
        }

        // #23: the offsets of a relative R1C1 reference are written with the invariant culture. In a culture whose minus sign is
        // U+2212 on .NET 5 and later (sv-SE, nb-NO, fi-FI) FormulaR1C1 returned R[−1]C[−5], which the library itself could not read back.
        [TestMethod]
        public void R1C1_references_do_not_depend_on_the_culture()
        {
            var culture = System.Threading.Thread.CurrentThread.CurrentCulture;
            try
            {
                foreach (string name in new[] { "sv-SE", "nb-NO", "fi-FI", "en-US", "ko-KR" })
                {
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo(name);
                    using (var p = new ExcelPackage())
                    {
                        var ws = p.Workbook.Worksheets.Add("S");
                        ws.Cells["G7"].Formula = "B6";
                        Assert.AreEqual("R[-1]C[-5]", ws.Cells["G7"].FormulaR1C1, name);
                        ws.Cells["D2"].Formula = "VLOOKUP(C2,A:B,1,0)";
                        Assert.AreEqual("VLOOKUP(RC[-1],C[-3]:C[-2],1,0)", ws.Cells["D2"].FormulaR1C1, name);

                        // what the library wrote, it reads back
                        ws.Cells["G9"].FormulaR1C1 = ws.Cells["G7"].FormulaR1C1;
                        Assert.AreEqual("B8", ws.Cells["G9"].Formula, name);

                        // a formula copied to other cells moves its references: left and up of the cell
                        ws.Cells["H10"].Formula = "SUM(F8:G9)+E10";
                        ws.Cells["H10"].Copy(ws.Cells["J12"]);
                        Assert.AreEqual("SUM(H10:I11)+G12", ws.Cells["J12"].Formula, name);
                        ws.Cells["H10"].Copy(ws.Cells["H13"]);
                        Assert.AreEqual("SUM(F11:G12)+E13", ws.Cells["H13"].Formula, name);
                        // a formula set on a range is a shared formula: each cell has its own references
                        ws.Cells["M10:M12"].Formula = "L9*2";
                        Assert.AreEqual("L11*2", ws.Cells["M12"].Formula, name);
                        Assert.AreEqual("R[-1]C[-1]*2", ws.Cells["M12"].FormulaR1C1, name);
                    }
                }
            }
            finally
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = culture;
            }
        }

        #region #21: the chart types of Excel 2016
        const string NsCx = "http://schemas.microsoft.com/office/drawing/2014/chartex";

        // A sheet with two category columns (A, B) and a value column (C) in rows 2 to 7, and a chart of the type on it
        static ExcelChartEx ExtendedChart(ExcelPackage a_package, eChartExType a_type, string a_sheet = "S")
        {
            var ws = a_package.Workbook.Worksheets.Add(a_sheet);
            ws.Cells["A1"].Value = "Grp"; ws.Cells["B1"].Value = "Name"; ws.Cells["C1"].Value = "Val";
            string[] names = { "a", "b", "c", "d", "e", "f" };
            for (int r = 2; r <= 7; r++) { ws.Cells[r, 1].Value = r <= 4 ? "g1" : "g2"; ws.Cells[r, 2].Value = names[r - 2]; ws.Cells[r, 3].Value = (r - 1) * 7; }
            return ws.Drawings.AddExtendedChart("Chart_" + a_type, a_type);
        }

        static XmlNamespaceManager CxNs(XmlElement a_root)
        {
            var ns = new XmlNamespaceManager(a_root.OwnerDocument.NameTable);
            ns.AddNamespace("cx", NsCx);
            ns.AddNamespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            ns.AddNamespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            ns.AddNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships");
            ns.AddNamespace("ct", "http://schemas.openxmlformats.org/package/2006/content-types");
            return ns;
        }

        // #21: a chart of Excel 2016 is a part of its own, with a style part and a colour part (Excel does not open the workbook
        // without them), and sits in the drawing inside an AlternateContent with a text box as the fallback
        [TestMethod]
        public void Extended_charts_are_written_with_their_parts()
        {
            foreach (eChartExType type in Enum.GetValues(typeof(eChartExType)))
            {
                using (var p = new ExcelPackage())
                {
                    ExcelChartEx chart = ExtendedChart(p, type);
                    bool hierarchy = type == eChartExType.Treemap || type == eChartExType.Sunburst;
                    var s = chart.Series.Add("C2:C7", hierarchy ? "A2:B7" : "B2:B7");
                    s.HeaderAddress = chart.WorkSheet.Cells["C1"];
                    chart.Title = "Title <" + type + ">";
                    chart.SetPosition(2, 0, 5, 0);
                    chart.SetSize(400, 250);
                    Assert.AreEqual("Chart_" + type, chart.Name);
                    Assert.AreEqual(type, chart.ChartType);

                    XmlElement types = SavedSheet(p, "[Content_Types].xml");
                    var ns = CxNs(types);
                    Assert.AreEqual("application/vnd.ms-office.chartex+xml", types.SelectSingleNode("ct:Override[@PartName='/xl/charts/chartEx1.xml']/@ContentType", ns).Value, type.ToString());
                    Assert.AreEqual("application/vnd.ms-office.chartstyle+xml", types.SelectSingleNode("ct:Override[@PartName='/xl/charts/style1.xml']/@ContentType", ns).Value, type.ToString());
                    Assert.AreEqual("application/vnd.ms-office.chartcolorstyle+xml", types.SelectSingleNode("ct:Override[@PartName='/xl/charts/colors1.xml']/@ContentType", ns).Value, type.ToString());

                    XmlElement drawingRels = SavedSheet(p, "xl/drawings/_rels/drawing1.xml.rels");
                    XmlNode rel = drawingRels.SelectSingleNode("rel:Relationship", CxNs(drawingRels));
                    Assert.AreEqual("http://schemas.microsoft.com/office/2014/relationships/chartEx", rel.Attributes["Type"].Value, type.ToString());
                    Assert.AreEqual("../charts/chartEx1.xml", rel.Attributes["Target"].Value, type.ToString());
                    XmlElement chartRels = SavedSheet(p, "xl/charts/_rels/chartEx1.xml.rels");
                    string[] targets = chartRels.SelectNodes("rel:Relationship", CxNs(chartRels)).OfType<XmlNode>().Select(n => n.Attributes["Type"].Value.Split('/').Last() + "=" + n.Attributes["Target"].Value).OrderBy(x => x).ToArray();
                    CollectionAssert.AreEqual(new[] { "chartColorStyle=colors1.xml", "chartStyle=style1.xml" }, targets, type.ToString());

                    XmlElement drawing = SavedSheet(p, "xl/drawings/drawing1.xml");
                    var dns = CxNs(drawing);
                    XmlNode anchor = drawing.SelectSingleNode("xdr:twoCellAnchor", dns);
                    Assert.AreEqual("from,to,AlternateContent,clientData", string.Join(",", anchor.ChildNodes.OfType<XmlElement>().Select(e => e.LocalName)), type.ToString());
                    XmlElement choice = (XmlElement)anchor.SelectSingleNode("mc:AlternateContent/mc:Choice", dns);
                    // funnel charts came one release later than the others and ask for the second namespace
                    string requires = type == eChartExType.Funnel ? "cx2" : "cx1";
                    Assert.AreEqual(requires, choice.GetAttribute("Requires"), type.ToString());
                    Assert.AreEqual(type == eChartExType.Funnel ? "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" : "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex", choice.GetAttribute("xmlns:" + requires), type.ToString());
                    Assert.AreEqual(rel.Attributes["Id"].Value, choice.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/cx:chart/@r:id", dns).Value, type.ToString());
                    Assert.AreEqual(NsCx, choice.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/@uri", dns).Value, type.ToString());
                    Assert.IsNotNull(anchor.SelectSingleNode("mc:AlternateContent/mc:Fallback/xdr:sp/xdr:txBody", dns), type.ToString());
                    Assert.AreEqual("5", anchor.SelectSingleNode("xdr:from/xdr:col", dns).InnerText, type.ToString());

                    XmlElement space = SavedSheet(p, "xl/charts/chartEx1.xml");
                    var cns = CxNs(space);
                    Assert.AreEqual(hierarchy ? "'S'!$A$2:$B$7" : "'S'!$B$2:$B$7", space.SelectSingleNode("cx:chartData/cx:data/cx:strDim[@type='cat']/cx:f", cns).InnerText, type.ToString());
                    Assert.AreEqual("'S'!$C$2:$C$7", space.SelectSingleNode("cx:chartData/cx:data/cx:numDim[@type='" + (hierarchy ? "size" : "val") + "']/cx:f", cns).InnerText, type.ToString());
                    Assert.AreEqual("Title <" + type + ">", space.SelectSingleNode("cx:chart/cx:title/cx:tx/cx:txData/cx:v", cns).InnerText, type.ToString());
                    Assert.AreEqual("'S'!$C$1", space.SelectSingleNode("cx:chart/cx:plotArea/cx:plotAreaRegion/cx:series/cx:tx/cx:txData/cx:f", cns).InnerText, type.ToString());
                    Assert.AreEqual("Val", space.SelectSingleNode("cx:chart/cx:plotArea/cx:plotAreaRegion/cx:series/cx:tx/cx:txData/cx:v", cns).InnerText, type.ToString());

                    string layouts = string.Join(",", space.SelectNodes("cx:chart/cx:plotArea/cx:plotAreaRegion/cx:series/@layoutId", cns).OfType<XmlNode>().Select(n => n.Value));
                    int axes = space.SelectNodes("cx:chart/cx:plotArea/cx:axis", cns).Count;
                    switch (type)
                    {
                        case eChartExType.Waterfall: Assert.AreEqual("waterfall", layouts); Assert.AreEqual(2, axes); break;
                        case eChartExType.Treemap: Assert.AreEqual("treemap", layouts); Assert.AreEqual(0, axes); break;
                        case eChartExType.Sunburst: Assert.AreEqual("sunburst", layouts); Assert.AreEqual(0, axes); break;
                        case eChartExType.Histogram: Assert.AreEqual("clusteredColumn", layouts); Assert.AreEqual(2, axes); break;
                        case eChartExType.Pareto: Assert.AreEqual("clusteredColumn,paretoLine", layouts); Assert.AreEqual(3, axes); break;
                        case eChartExType.BoxWhisker: Assert.AreEqual("boxWhisker", layouts); Assert.AreEqual(2, axes); break;
                        case eChartExType.Funnel: Assert.AreEqual("funnel", layouts); Assert.AreEqual(1, axes); break;
                    }

                    // the style part has an entry for every element of a chart, each with its four references
                    XmlElement style = SavedSheet(p, "xl/charts/style1.xml");
                    var sns = new XmlNamespaceManager(style.OwnerDocument.NameTable);
                    sns.AddNamespace("cs", "http://schemas.microsoft.com/office/drawing/2012/chartStyle");
                    Assert.AreEqual(31, style.ChildNodes.OfType<XmlElement>().Count(), type.ToString());
                    foreach (XmlElement entry in style.ChildNodes.OfType<XmlElement>().Where(e => e.LocalName != "dataPointMarkerLayout"))
                    {
                        Assert.AreEqual("lnRef,fillRef,effectRef,fontRef", string.Join(",", entry.ChildNodes.OfType<XmlElement>().Take(4).Select(e => e.LocalName)), type + " " + entry.LocalName);
                    }
                    Assert.AreEqual("colorStyle", SavedSheet(p, "xl/charts/colors1.xml").LocalName);
                }
            }
        }

        // #21: what belongs to one chart type, and the things a chart can be without
        [TestMethod]
        public void Extended_charts_write_the_options_of_their_type()
        {
            using (var p = new ExcelPackage())
            {
                string region = "cx:chart/cx:plotArea/cx:plotAreaRegion/";

                // waterfall: the totals; no title, no labels, no legend
                ExcelChartEx waterfall = ExtendedChart(p, eChartExType.Waterfall, "Waterfall");
                Assert.IsTrue(waterfall.DataLabels && waterfall.ShowLegend, "on when the chart is made, as in Excel");
                waterfall.Series.Add("C2:C7", "B2:B7").Subtotals = new[] { 0, 5 };
                waterfall.DataLabels = false; waterfall.ShowLegend = false;
                XmlElement w = SavedSheet(p, "xl/charts/chartEx1.xml");
                var ns = CxNs(w);
                Assert.AreEqual("0,5", string.Join(",", w.SelectNodes(region + "cx:series/cx:layoutPr/cx:subtotals/cx:idx/@val", ns).OfType<XmlNode>().Select(n => n.Value)));
                Assert.IsNull(w.SelectSingleNode("cx:chart/cx:title", ns));
                Assert.IsNull(w.SelectSingleNode("cx:chart/cx:legend", ns));
                Assert.IsNull(w.SelectSingleNode(region + "cx:series/cx:dataLabels", ns));
                Assert.IsNull(w.SelectSingleNode(region + "cx:series/cx:tx", ns), "no name was given");
                Assert.AreEqual("dataId,layoutPr", string.Join(",", ChildNames((XmlElement)w.SelectSingleNode(region + "cx:series", ns))));

                // histogram: of the values alone, with the width of a bin, the number of bins, or neither
                ExcelChartEx histogram = ExtendedChart(p, eChartExType.Histogram, "Histogram");
                histogram.Series.Add("C2:C7", null).Header = "Elapsed & more";
                histogram.BinWidth = 2.5;
                XmlElement h = SavedSheet(p, "xl/charts/chartEx2.xml");
                Assert.IsNull(h.SelectSingleNode("cx:chartData/cx:data/cx:strDim", ns));
                Assert.AreEqual("'Histogram'!$C$2:$C$7", h.SelectSingleNode("cx:chartData/cx:data/cx:numDim/cx:f", ns).InnerText);
                Assert.AreEqual("2.5", h.SelectSingleNode(region + "cx:series/cx:layoutPr/cx:binning[@intervalClosed='r']/cx:binSize/@val", ns).Value);
                Assert.AreEqual("Elapsed & more", h.SelectSingleNode(region + "cx:series/cx:tx/cx:txData/cx:v", ns).InnerText);
                Assert.IsNull(h.SelectSingleNode(region + "cx:series/cx:tx/cx:txData/cx:f", ns));
                histogram.BinWidth = null; histogram.BinCount = 4;
                h = SavedSheet(p, "xl/charts/chartEx2.xml");
                Assert.AreEqual("4", h.SelectSingleNode(region + "cx:series/cx:layoutPr/cx:binning/cx:binCount/@val", ns).Value);
                histogram.BinCount = null;
                h = SavedSheet(p, "xl/charts/chartEx2.xml");
                Assert.AreEqual(0, h.SelectSingleNode(region + "cx:series/cx:layoutPr/cx:binning", ns).ChildNodes.Count);

                // Pareto: of categories the columns are the categories, of values alone they are bins
                ExcelChartEx pareto = ExtendedChart(p, eChartExType.Pareto, "Pareto");
                pareto.Series.Add("C2:C7", "B2:B7");
                XmlElement pa = SavedSheet(p, "xl/charts/chartEx3.xml");
                Assert.IsNotNull(pa.SelectSingleNode(region + "cx:series[@layoutId='clusteredColumn']/cx:layoutPr/cx:aggregation", ns));
                Assert.AreEqual("0", pa.SelectSingleNode(region + "cx:series[@layoutId='paretoLine']/@ownerIdx", ns).Value);
                Assert.AreEqual("percentage", pa.SelectSingleNode("cx:chart/cx:plotArea/cx:axis[@id='2']/cx:units/@unit", ns).Value);

                // an address with a sheet name, a sheet name that needs quotes, ranges of another sheet
                var other = p.Workbook.Worksheets.Add("It's data");
                other.Cells["A1"].Value = "x"; other.Cells["B1"].Value = 1;
                ExcelChartEx funnel = ExtendedChart(p, eChartExType.Funnel, "Funnel");
                funnel.Series.Add("'It''s data'!B1:B5", "'It''s data'!A1:A5");
                funnel.Series.Add(other.Cells["B1:B5"], null);
                XmlElement f = SavedSheet(p, "xl/charts/chartEx4.xml");
                Assert.AreEqual("'It''s data'!$B$1:$B$5", f.SelectNodes("cx:chartData/cx:data/cx:numDim/cx:f", ns)[0].InnerText);
                Assert.AreEqual("'It''s data'!$B$1:$B$5", f.SelectNodes("cx:chartData/cx:data/cx:numDim/cx:f", ns)[1].InnerText);
                Assert.AreEqual("0,1", string.Join(",", f.SelectNodes(region + "cx:series/cx:dataId/@val", ns).OfType<XmlNode>().Select(n => n.Value)));
                Assert.AreEqual(2, funnel.Series.Count);

                Assert.ThrowsException<ArgumentException>(() => funnel.Series.Add("", "A1:A5"));
                Assert.ThrowsException<Exception>(() => p.Workbook.Worksheets["Funnel"].Drawings.AddExtendedChart("Chart_Funnel", eChartExType.Funnel), "a name that is taken");
                var chartSheet = p.Workbook.Worksheets.AddChart("ChartSheet", eChartType.Line);
                Assert.ThrowsException<NotSupportedException>(() => chartSheet.Drawings.AddExtendedChart("X", eChartExType.Funnel));
            }
        }

        // #21: a workbook with such a chart keeps it when it is loaded and saved; removing the chart removes its parts
        [TestMethod]
        public void Extended_charts_survive_a_load_and_are_removed_with_their_parts()
        {
            using (var p = new ExcelPackage())
            {
                ExcelChartEx chart = ExtendedChart(p, eChartExType.Treemap);
                chart.Series.Add("C2:C7", "A2:B7");
                chart.Title = "Sizes";
                ExcelChartEx second = p.Workbook.Worksheets["S"].Drawings.AddExtendedChart("Second", eChartExType.Histogram);
                second.Series.Add("C2:C7", null);
                var column = p.Workbook.Worksheets["S"].Drawings.AddChart("Column", eChartType.ColumnClustered);
                column.Series.Add("C2:C7", "B2:B7");
                string before = SavedSheet(p, "xl/charts/chartEx1.xml").OuterXml;

                using (var r = Reload(p))
                {
                    var ws = r.Workbook.Worksheets["S"];
                    Assert.AreEqual(3, ws.Drawings.Count);
                    ws.Cells["C2"].Value = 99;
                    Assert.AreEqual(before, SavedSheet(r, "xl/charts/chartEx1.xml").OuterXml, "the chart part is kept as it is");
                    Assert.IsNotNull(SavedSheet(r, "xl/charts/style2.xml"));
                    Assert.IsNotNull(SavedSheet(r, "xl/charts/_rels/chartEx2.xml.rels"));
                    XmlElement drawing = SavedSheet(r, "xl/drawings/drawing1.xml");
                    Assert.AreEqual(2, drawing.SelectNodes("xdr:twoCellAnchor/mc:AlternateContent", CxNs(drawing)).Count);
                }

                // removed before the save: the three parts of the chart go with it, the other chart stays
                p.Workbook.Worksheets["S"].Drawings.Remove(chart);
                Assert.AreEqual(2, p.Workbook.Worksheets["S"].Drawings.Count);
                var ms = new MemoryStream();
                p.SaveAs(ms);
                ms.Position = 0;
                using (var zip = new System.IO.Compression.ZipArchive(ms))
                {
                    string[] parts = zip.Entries.Select(e => e.FullName).Where(n => n.StartsWith("xl/charts/")).OrderBy(n => n).ToArray();
                    CollectionAssert.AreEqual(new[] { "xl/charts/_rels/chartEx2.xml.rels", "xl/charts/chart1.xml", "xl/charts/chartEx2.xml", "xl/charts/colors2.xml", "xl/charts/style2.xml" }, parts, string.Join(" ", parts));
                }
            }
        }

        // #21: a copy of a worksheet has a copy of the chart with parts of its own. The drawing of the copy pointed at a
        // relationship that did not exist, and Excel did not open the workbook.
        [TestMethod]
        public void Extended_charts_are_copied_with_the_worksheet()
        {
            using (var p = new ExcelPackage())
            {
                ExcelChartEx chart = ExtendedChart(p, eChartExType.Funnel);
                chart.Series.Add("C2:C7", "B2:B7");
                chart.Title = "Before the copy";
                var copy = p.Workbook.Worksheets.Add("Copy", p.Workbook.Worksheets["S"]);
                chart.Title = "After the copy";
                Assert.AreEqual(1, copy.Drawings.Count);

                // loaded and copied once more: the chart is then a drawing the library only keeps
                using (var r = Reload(p))
                {
                    r.Workbook.Worksheets.Add("Copy2", r.Workbook.Worksheets["Copy"]);
                    var ms = new MemoryStream();
                    r.SaveAs(ms);
                    ms.Position = 0;
                    using (var zip = new System.IO.Compression.ZipArchive(ms))
                    {
                        string[] names = zip.Entries.Select(e => e.FullName).ToArray();
                        Assert.AreEqual(3, names.Count(n => System.Text.RegularExpressions.Regex.IsMatch(n, @"^xl/charts/chartEx\d+\.xml$")), string.Join(" ", names));
                        Assert.AreEqual(3, names.Count(n => System.Text.RegularExpressions.Regex.IsMatch(n, @"^xl/charts/style\d+\.xml$")), string.Join(" ", names));
                        Assert.AreEqual(3, names.Count(n => System.Text.RegularExpressions.Regex.IsMatch(n, @"^xl/charts/colors\d+\.xml$")), string.Join(" ", names));

                        // every drawing points at a chart part of its own that is in the package, and that part at its style and colours
                        var seen = new System.Collections.Generic.HashSet<string>();
                        foreach (string drawingName in names.Where(n => System.Text.RegularExpressions.Regex.IsMatch(n, @"^xl/drawings/drawing\d+\.xml$")))
                        {
                            var drawing = new XmlDocument(); using (var s = zip.GetEntry(drawingName).Open()) drawing.Load(s);
                            var ns = CxNs(drawing.DocumentElement);
                            string id = drawing.SelectSingleNode("//cx:chart/@r:id", ns).Value;
                            var rels = new XmlDocument(); using (var s = zip.GetEntry(drawingName.Replace("drawings/", "drawings/_rels/") + ".rels").Open()) rels.Load(s);
                            string target = rels.SelectSingleNode("//rel:Relationship[@Id='" + id + "']/@Target", ns).Value;
                            string chartName = "xl/charts/" + target.Split('/').Last();
                            Assert.IsTrue(names.Contains(chartName), drawingName + " -> " + target);
                            Assert.IsTrue(seen.Add(chartName), "two drawings share " + chartName);
                            var chartRels = new XmlDocument(); using (var s = zip.GetEntry(chartName.Replace("charts/", "charts/_rels/") + ".rels").Open()) chartRels.Load(s);
                            foreach (XmlNode t in chartRels.SelectNodes("//rel:Relationship/@Target", ns)) Assert.IsTrue(names.Contains("xl/charts/" + t.Value), chartName + " -> " + t.Value);
                            Assert.AreEqual(2, chartRels.SelectNodes("//rel:Relationship", ns).Count, chartName);
                        }
                        Assert.AreEqual(3, seen.Count);

                        // the copy was made before the title changed: it has the title of that moment
                        string titles = string.Join("|", seen.OrderBy(n => n).Select(n => { var d = new XmlDocument(); using (var s = zip.GetEntry(n).Open()) d.Load(s); return d.SelectSingleNode("//cx:title//cx:v", CxNs(d.DocumentElement)).InnerText; }));
                        Assert.AreEqual("After the copy|Before the copy|Before the copy", titles);
                    }
                }
            }
        }
        #endregion

        #region #19: the format of chart series
        static readonly string[] s_serOrder = { "idx", "order", "tx", "spPr", "invertIfNegative", "pictureOptions", "marker", "explosion", "dPt", "dLbls", "trendline", "errBars", "cat", "val", "xVal", "yVal", "smooth", "shape", "bubbleSize", "bubble3D", "extLst" };
        static readonly System.Drawing.Color s_red = System.Drawing.Color.FromArgb(0xC0, 0x00, 0x00);
        static readonly System.Drawing.Color s_blue = System.Drawing.Color.FromArgb(0x00, 0x70, 0xC0);
        static readonly System.Drawing.Color s_green = System.Drawing.Color.FromArgb(0x00, 0xB0, 0x50);

        // A chart of the given type on a sheet with a category column and four value columns, with its series
        static ExcelChart FormatChart(ExcelPackage a_package, eChartType a_type, string a_sheet = "S")
        {
            var ws = a_package.Workbook.Worksheets.Add(a_sheet);
            for (int r = 1; r <= 6; r++) for (int c = 1; c <= 5; c++) ws.Cells[r, c].Value = r == 1 ? (object)("H" + c) : r * c + 1;
            var chart = ws.Drawings.AddChart("C_" + a_sheet, a_type);
            int series = a_type == eChartType.StockHLC ? 3 : 2;
            for (int c = 2; c <= series + 1; c++)
            {
                if (a_type == eChartType.Bubble) ((ExcelBubbleChartSeries)chart.Series).Add(ExcelCellBase.GetAddress(2, c, 6, c), "A2:A6", "E2:E6");
                else chart.Series.Add(ExcelCellBase.GetAddress(2, c, 6, c), "A2:A6");
            }
            return chart;
        }

        static XmlNamespaceManager ChartNs(XmlElement a_chartSpace)
        {
            var ns = new XmlNamespaceManager(a_chartSpace.OwnerDocument.NameTable);
            ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            return ns;
        }

        // Every series has its children in the order of the schema, and no shape or line has more than one fill
        static void AssertSeriesAreValid(XmlElement a_chartSpace, string a_what)
        {
            var ns = ChartNs(a_chartSpace);
            foreach (XmlNode ser in a_chartSpace.SelectNodes("//c:ser", ns))
            {
                string[] names = ser.ChildNodes.OfType<XmlElement>().Select(e => e.LocalName).ToArray();
                int[] positions = names.Select(n => Array.IndexOf(s_serOrder, n)).ToArray();
                Assert.IsTrue(positions.All(p => p >= 0), a_what + ": " + string.Join(",", names));
                CollectionAssert.AreEqual(positions.OrderBy(p => p).ToArray(), positions, a_what + ": " + string.Join(",", names));
            }
            foreach (XmlNode shape in a_chartSpace.SelectNodes("//c:spPr|//a:ln", ns))
            {
                int fills = shape.ChildNodes.OfType<XmlElement>().Count(e => e.LocalName.EndsWith("Fill"));
                Assert.IsTrue(fills <= 1, a_what + ": " + shape.OuterXml);
            }
        }

        // #19: the fill and the border of a series: after c:tx, with one fill per shape and per line, on every chart type.
        // On a scatter and a stock chart the solid fill was added next to the noFill of the line, and Excel did not open the file.
        [TestMethod]
        public void Series_fill_and_border_are_valid_on_every_chart_type()
        {
            var types = new[]
            {
                eChartType.ColumnClustered, eChartType.BarStacked, eChartType.ColumnClustered3D, eChartType.Line, eChartType.LineMarkers, eChartType.Pie,
                eChartType.Doughnut, eChartType.Area, eChartType.XYScatter, eChartType.XYScatterLines, eChartType.RadarMarkers, eChartType.Bubble, eChartType.StockHLC
            };
            foreach (var type in types)
            {
                using (var p = new ExcelPackage())
                {
                    var chart = FormatChart(p, type);
                    chart.Series[0].Fill.Color = s_red;
                    chart.Series[0].Border.Fill.Color = s_blue;
                    chart.Series[0].Border.Width = 2;
                    chart.Series[1].Border.Fill.Color = s_green;   // a border without a fill

                    XmlElement space = SavedSheet(p, "xl/charts/chart1.xml");
                    var ns = ChartNs(space);
                    AssertSeriesAreValid(space, type.ToString());
                    XmlNode first = space.SelectNodes("//c:ser", ns)[0];
                    Assert.AreEqual("C00000", first.SelectSingleNode("c:spPr/a:solidFill/a:srgbClr/@val", ns).Value, type.ToString());
                    Assert.AreEqual("0070C0", first.SelectSingleNode("c:spPr/a:ln/a:solidFill/a:srgbClr/@val", ns).Value, type.ToString());
                    Assert.AreEqual("00B050", space.SelectNodes("//c:ser", ns)[1].SelectSingleNode("c:spPr/a:ln/a:solidFill/a:srgbClr/@val", ns).Value, type.ToString());
                    Assert.AreEqual("spPr", ((XmlElement)first.SelectSingleNode("c:tx", ns).NextSibling).LocalName, type.ToString());

                    using (var r = Reload(p))
                    {
                        var rc = (ExcelChart)r.Workbook.Worksheets["S"].Drawings[0];
                        Assert.AreEqual(s_red.ToArgb() & 0xFFFFFF, rc.Series[0].Fill.Color.ToArgb() & 0xFFFFFF, type.ToString());
                        Assert.AreEqual(s_blue.ToArgb() & 0xFFFFFF, rc.Series[0].Border.Fill.Color.ToArgb() & 0xFFFFFF, type.ToString());
                    }
                }
            }

            // no fill after a colour, and a colour after no fill: one of them is left
            using (var p = new ExcelPackage())
            {
                var chart = FormatChart(p, eChartType.ColumnClustered);
                chart.Series[0].Fill.Color = s_red;
                chart.Series[0].Fill.Style = eFillStyle.NoFill;
                chart.Series[1].Fill.Style = eFillStyle.NoFill;
                chart.Series[1].Fill.Style = eFillStyle.SolidFill;
                chart.Series[1].Fill.Color = s_green;
                XmlElement space = SavedSheet(p, "xl/charts/chart1.xml");
                var ns = ChartNs(space);
                AssertSeriesAreValid(space, "no fill");
                Assert.AreEqual("<a:noFill xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" />", space.SelectNodes("//c:ser", ns)[0].SelectSingleNode("c:spPr", ns).InnerXml);
                Assert.AreEqual("00B050", space.SelectNodes("//c:ser", ns)[1].SelectSingleNode("c:spPr/a:solidFill/a:srgbClr/@val", ns).Value);
            }

            // the line colour of a scatter series without lines: the same line, through the property of the scatter series
            using (var p = new ExcelPackage())
            {
                var chart = FormatChart(p, eChartType.XYScatter);
                ((ExcelScatterChartSerie)chart.Series[0]).LineColor = s_blue;
                ((ExcelScatterChartSerie)chart.Series[0]).MarkerColor = s_red;
                XmlElement space = SavedSheet(p, "xl/charts/chart1.xml");
                var ns = ChartNs(space);
                AssertSeriesAreValid(space, "scatter line colour");
                Assert.AreEqual("solidFill", string.Join(",", ChildNames((XmlElement)space.SelectNodes("//c:ser", ns)[0].SelectSingleNode("c:spPr/a:ln", ns))));
                Assert.AreEqual("noFill", string.Join(",", ChildNames((XmlElement)space.SelectNodes("//c:ser", ns)[1].SelectSingleNode("c:spPr/a:ln", ns))), "the other series keeps its line without a fill");
            }

            // the other users of the fill class: the chart area, the plot area and a shape keep one fill too
            using (var p = new ExcelPackage())
            {
                var chart = FormatChart(p, eChartType.ColumnClustered);
                chart.Fill.Style = eFillStyle.NoFill;
                chart.Fill.Color = s_red;
                chart.Border.Fill.Style = eFillStyle.NoFill;
                chart.Border.Fill.Color = s_blue;
                chart.PlotArea.Fill.Color = s_green;
                chart.PlotArea.Fill.Style = eFillStyle.NoFill;
                var shape = chart.WorkSheet.Drawings.AddShape("Shape", eShapeStyle.Rect);
                shape.Fill.Style = eFillStyle.NoFill;
                shape.Fill.Color = s_green;
                shape.Border.Fill.Style = eFillStyle.NoFill;
                shape.Border.Fill.Color = s_blue;

                XmlElement space = SavedSheet(p, "xl/charts/chart1.xml");
                var ns = ChartNs(space);
                AssertSeriesAreValid(space, "chart and plot area");
                Assert.AreEqual("C00000", space.SelectSingleNode("c:spPr/a:solidFill/a:srgbClr/@val", ns).Value);
                Assert.AreEqual("0070C0", space.SelectSingleNode("c:spPr/a:ln/a:solidFill/a:srgbClr/@val", ns).Value);
                Assert.IsNotNull(space.SelectSingleNode("c:chart/c:plotArea/c:spPr/a:noFill", ns));

                XmlElement drawing = SavedSheet(p, "xl/drawings/drawing1.xml");
                var dns = new XmlNamespaceManager(drawing.OwnerDocument.NameTable);
                dns.AddNamespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
                dns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
                XmlNode sp = drawing.SelectSingleNode("//xdr:sp/xdr:spPr", dns);
                Assert.AreEqual(1, sp.ChildNodes.OfType<XmlElement>().Count(e => e.LocalName.EndsWith("Fill")), sp.OuterXml);
                Assert.AreEqual(1, sp.SelectSingleNode("a:ln", dns).ChildNodes.OfType<XmlElement>().Count(e => e.LocalName.EndsWith("Fill")), sp.OuterXml);
                Assert.AreEqual("00B050", sp.SelectSingleNode("a:solidFill/a:srgbClr/@val", dns).Value);
            }
        }

        // #19: what a series could have before (markers, line colour, explosion, data labels, trendlines) is written in the order of the schema too
        [TestMethod]
        public void Series_children_are_in_the_order_of_the_schema()
        {
            using (var p = new ExcelPackage())
            {
                var bar = FormatChart(p, eChartType.ColumnClustered, "Bar");
                bar.Series[0].TrendLines.Add(eTrendLine.Linear);
                ((ExcelBarChartSerie)bar.Series[0]).DataLabel.ShowValue = true;
                bar.Series[0].Fill.Color = s_red;
                bar.Series[0].DataPoints.Add(1).Fill.Color = s_green;
                bar.Series[0].AddErrorBars(eErrorBarType.Both, eErrorValueType.FixedValue, 1);
                Assert.AreEqual("idx,order,tx,spPr,invertIfNegative,dPt,dLbls,trendline,errBars,cat,val",
                    string.Join(",", ChildNames((XmlElement)SavedSheet(p, "xl/charts/chart1.xml").SelectSingleNode("//c:ser", ChartNs(SavedSheet(p, "xl/charts/chart1.xml"))))));

                var line = FormatChart(p, eChartType.LineMarkers, "Line");
                var ls = (ExcelLineChartSerie)line.Series[0];
                ls.DataLabel.ShowValue = true;
                ls.TrendLines.Add(eTrendLine.MovingAvgerage);
                ls.LineColor = s_blue; ls.Marker = eMarkerStyle.Circle; ls.MarkerSize = 9; ls.MarkerLineColor = s_red; ls.Smooth = true;
                ls.DataPoints.Add(0).MarkerFill.Color = s_green;

                var pie = FormatChart(p, eChartType.Pie, "Pie");
                var ps = (ExcelPieChartSerie)pie.Series[0];
                ps.DataPoints.Add(2).Fill.Color = s_green;
                ps.DataLabel.ShowPercent = true;
                ps.Explosion = 10;
                ps.Fill.Color = s_red;

                var scatter = FormatChart(p, eChartType.XYScatterLines, "Scatter");
                var ss = (ExcelScatterChartSerie)scatter.Series[0];
                ss.DataLabel.ShowValue = true;
                ss.Marker = eMarkerStyle.Diamond; ss.MarkerSize = 8; ss.MarkerColor = s_red; ss.LineColor = s_blue;
                ss.AddErrorBars(eErrorBarType.Both, eErrorValueType.Percentage, 5);

                for (int i = 1; i <= 4; i++) AssertSeriesAreValid(SavedSheet(p, "xl/charts/chart" + i + ".xml"), "chart" + i);
                XmlElement lineSpace = SavedSheet(p, "xl/charts/chart2.xml");
                Assert.AreEqual("symbol,size,spPr", string.Join(",", ChildNames((XmlElement)lineSpace.SelectSingleNode("//c:ser/c:marker", ChartNs(lineSpace)))));
                XmlElement pieSpace = SavedSheet(p, "xl/charts/chart3.xml");
                Assert.AreEqual("idx,order,tx,spPr,explosion,dPt,dLbls,cat,val", string.Join(",", ChildNames((XmlElement)pieSpace.SelectSingleNode("//c:ser", ChartNs(pieSpace)))));
            }
        }

        // #19: single points of a series with a format of their own
        [TestMethod]
        public void Data_points_are_written_as_Excel_writes_them()
        {
            using (var p = new ExcelPackage())
            {
                var bar = FormatChart(p, eChartType.ColumnClustered, "Bar");
                var s = bar.Series[0];
                Assert.AreEqual(0, s.DataPoints.Count);
                Assert.IsNull(s.DataPoints[2]);
                ExcelChartDataPoint p2 = s.DataPoints.Add(2);
                p2.Fill.Color = s_green;
                s.DataPoints.Add(4).Border.Fill.Color = s_blue;
                s.DataPoints.Add(0).Fill.Color = s_red;
                Assert.AreSame(p2, s.DataPoints.Add(2), "the point that is there");
                Assert.AreSame(p2, s.DataPoints[2]);
                Assert.AreEqual("0,2,4", string.Join(",", s.DataPoints.Select(x => x.Index)));

                var line = FormatChart(p, eChartType.Line, "Line");
                line.Series[0].DataPoints.Add(1).MarkerFill.Color = s_green;
                line.Series[0].DataPoints[1].MarkerBorder.Fill.Color = s_blue;
                line.Series[0].DataPoints[1].Border.Fill.Color = s_red;

                XmlElement space = SavedSheet(p, "xl/charts/chart1.xml");
                var ns = ChartNs(space);
                AssertSeriesAreValid(space, "bar");
                XmlNodeList points = space.SelectNodes("//c:ser[1]/c:dPt", ns);
                Assert.AreEqual("0,2,4", string.Join(",", points.OfType<XmlNode>().Select(n => n.SelectSingleNode("c:idx/@val", ns).Value)));
                Assert.AreEqual("<c:idx val=\"2\" /><c:invertIfNegative val=\"0\" /><c:bubble3D val=\"0\" /><c:spPr><a:solidFill><a:srgbClr val=\"00B050\" /></a:solidFill></c:spPr>",
                    points[1].InnerXml.Replace(" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"", "").Replace(" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"", ""));
                XmlElement lineSpace = SavedSheet(p, "xl/charts/chart2.xml");
                AssertSeriesAreValid(lineSpace, "line");
                Assert.AreEqual("<c:idx val=\"1\" /><c:marker><c:spPr><a:solidFill><a:srgbClr val=\"00B050\" /></a:solidFill><a:ln><a:solidFill><a:srgbClr val=\"0070C0\" /></a:solidFill></a:ln></c:spPr></c:marker><c:bubble3D val=\"0\" /><c:spPr><a:ln><a:solidFill><a:srgbClr val=\"C00000\" /></a:solidFill></a:ln></c:spPr>",
                    lineSpace.SelectSingleNode("//c:ser[1]/c:dPt", ChartNs(lineSpace)).InnerXml.Replace(" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"", "").Replace(" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"", ""));

                using (var r = Reload(p))
                {
                    var rs = ((ExcelChart)r.Workbook.Worksheets["Bar"].Drawings[0]).Series[0];
                    Assert.AreEqual(3, rs.DataPoints.Count);
                    Assert.AreEqual(s_green.ToArgb() & 0xFFFFFF, rs.DataPoints[2].Fill.Color.ToArgb() & 0xFFFFFF);
                    Assert.IsTrue(rs.DataPoints.Remove(2));
                    Assert.IsFalse(rs.DataPoints.Remove(2));
                    rs.DataPoints.Add(3).Fill.Color = s_green;
                    XmlElement again = SavedSheet(r, "xl/charts/chart1.xml");
                    Assert.AreEqual("0,3,4", string.Join(",", again.SelectNodes("//c:ser[1]/c:dPt/c:idx/@val", ChartNs(again)).OfType<XmlNode>().Select(n => n.Value)));
                    AssertSeriesAreValid(again, "after a load");
                }

                var surface = FormatChart(p, eChartType.Surface, "Surface");
                Assert.ThrowsException<NotSupportedException>(() => surface.Series[0].DataPoints.Add(0));
                Assert.ThrowsException<ArgumentOutOfRangeException>(() => bar.Series[0].DataPoints.Add(-1));
            }
        }

        // #19: error bars, for the chart types that have them
        [TestMethod]
        public void Error_bars_are_written_as_Excel_writes_them()
        {
            string c = " xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"";
            using (var p = new ExcelPackage())
            {
                var bar = FormatChart(p, eChartType.ColumnClustered, "Bar");
                Assert.IsNull(bar.Series[0].ErrorBars);
                bar.Series[0].AddErrorBars(eErrorBarType.Plus, eErrorValueType.Percentage, 5);
                var e = bar.Series[0].AddErrorBars(eErrorBarType.Both, eErrorValueType.FixedValue, 1.5);   // in place of the first
                bar.Series[1].AddErrorBars(eErrorBarType.Minus, eErrorValueType.StandardError);
                Assert.AreEqual(eErrorBarType.Both, bar.Series[0].ErrorBars.Type);
                Assert.AreEqual(1.5, bar.Series[0].ErrorBars.Value);

                var scatter = FormatChart(p, eChartType.XYScatter, "Scatter");
                scatter.Series[0].AddErrorBars(eErrorBarType.Both, eErrorValueType.StandardDeviation, 2);            // y first
                scatter.Series[0].AddErrorBars(eErrorBarType.Plus, eErrorValueType.Percentage, 10, horizontal: true);  // then x: written before y
                scatter.Series[0].ErrorBars.Border.Fill.Color = s_blue;
                scatter.Series[0].ErrorBars.NoEndCap = true;

                var line = FormatChart(p, eChartType.Line, "Line");
                line.Series[0].AddErrorBars(eErrorBarType.Both, eErrorValueType.FixedValue, 3);
                line.Series[1].AddErrorBars(eErrorBarType.Both, eErrorValueType.FixedValue, 3);
                line.Series[1].RemoveErrorBars();

                XmlElement space = SavedSheet(p, "xl/charts/chart1.xml");
                var ns = ChartNs(space);
                AssertSeriesAreValid(space, "bar");
                Assert.AreEqual(1, space.SelectNodes("//c:ser[1]/c:errBars", ns).Count);
                Assert.AreEqual("<c:errBarType val=\"both\" /><c:errValType val=\"fixedVal\" /><c:noEndCap val=\"0\" /><c:val val=\"1.5\" />", space.SelectSingleNode("//c:ser[1]/c:errBars", ns).InnerXml.Replace(c, ""));
                Assert.AreEqual("<c:errBarType val=\"minus\" /><c:errValType val=\"stdErr\" /><c:noEndCap val=\"0\" />", space.SelectSingleNode("//c:ser[2]/c:errBars", ns).InnerXml.Replace(c, ""));

                XmlElement scatterSpace = SavedSheet(p, "xl/charts/chart2.xml");
                AssertSeriesAreValid(scatterSpace, "scatter");
                XmlNodeList bars = scatterSpace.SelectNodes("//c:ser[1]/c:errBars", ChartNs(scatterSpace));
                Assert.AreEqual(2, bars.Count);
                Assert.AreEqual("<c:errDir val=\"x\" /><c:errBarType val=\"plus\" /><c:errValType val=\"percentage\" /><c:noEndCap val=\"0\" /><c:val val=\"10\" />", bars[0].InnerXml.Replace(c, ""));
                Assert.AreEqual("<c:errDir val=\"y\" /><c:errBarType val=\"both\" /><c:errValType val=\"stdDev\" /><c:noEndCap val=\"1\" /><c:val val=\"2\" /><c:spPr><a:ln><a:solidFill><a:srgbClr val=\"0070C0\" /></a:solidFill></a:ln></c:spPr>",
                    bars[1].InnerXml.Replace(c, "").Replace(" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"", ""));

                XmlElement lineSpace = SavedSheet(p, "xl/charts/chart3.xml");
                Assert.AreEqual("y", lineSpace.SelectSingleNode("//c:ser[1]/c:errBars/c:errDir/@val", ChartNs(lineSpace)).Value);
                Assert.AreEqual(0, lineSpace.SelectNodes("//c:ser[2]/c:errBars", ChartNs(lineSpace)).Count);

                using (var r = Reload(p))
                {
                    var rs = ((ExcelChart)r.Workbook.Worksheets["Scatter"].Drawings[0]).Series[0];
                    Assert.AreEqual(eErrorValueType.StandardDeviation, rs.ErrorBars.ValueType);
                    Assert.IsTrue(rs.ErrorBars.NoEndCap);
                    Assert.IsFalse(rs.ErrorBars.Horizontal);
                    Assert.AreEqual(eErrorBarType.Plus, rs.ErrorBarsX.Type);
                    Assert.AreEqual(10d, rs.ErrorBarsX.Value);
                    Assert.IsNull(((ExcelChart)r.Workbook.Worksheets["Bar"].Drawings[0]).Series[0].ErrorBarsX);
                }

                var pie = FormatChart(p, eChartType.Pie, "Pie");
                Assert.ThrowsException<NotSupportedException>(() => pie.Series[0].AddErrorBars(eErrorBarType.Both, eErrorValueType.FixedValue, 1));
                Assert.ThrowsException<NotSupportedException>(() => bar.Series[0].AddErrorBars(eErrorBarType.Both, eErrorValueType.FixedValue, 1, horizontal: true));
                Assert.ThrowsException<ArgumentOutOfRangeException>(() => bar.Series[0].AddErrorBars(eErrorBarType.Both, eErrorValueType.FixedValue, -1));
            }
        }
        #endregion

        // The sheet of the filter tests: a header and ten rows (2 to 11) with a name, a number shown with two decimals,
        // a date, a note with empty cells in rows 5 and 9, and a group. The same data as in the file Excel filtered itself.
        static ExcelWorksheet FilterSheet(ExcelPackage a_package, string a_name = "S")
        {
            var ws = a_package.Workbook.Worksheets.Add(a_name);
            string[] heads = { "Name", "Qty", "When", "Note", "Grp" };
            string[] names = { "alpha", "beta", "gamma", "alpha", "delta", "beta", "Alpha", "epsilon", "gamma", "zeta" };
            for (int c = 0; c < heads.Length; c++) ws.Cells[1, c + 1].Value = heads[c];
            for (int i = 1; i <= 10; i++)
            {
                ws.Cells[i + 1, 1].Value = names[i - 1];
                ws.Cells[i + 1, 2].Value = i * 3;
                ws.Cells[i + 1, 3].Value = new DateTime(2026, 1, i);
                if (i % 4 != 0) ws.Cells[i + 1, 4].Value = "note " + i;
                ws.Cells[i + 1, 5].Value = "g" + (i % 3);
            }
            ws.Cells["B2:B11"].Style.Numberformat.Format = "0.00";
            ws.Cells["C2:C11"].Style.Numberformat.Format = "yyyy-mm-dd";
            return ws;
        }

        static string HiddenRows(ExcelWorksheet a_ws, int a_from = 2, int a_to = 11)
        {
            return string.Join(",", Enumerable.Range(a_from, a_to - a_from + 1).Where(r => a_ws.Row(r).Hidden));
        }

        static string InnerOf(XmlElement a_element)
        {
            return a_element.InnerXml.Replace(" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"", "");
        }

        // #17: the filters of a worksheet autofilter are written as Excel writes them, and the rows that do not pass are hidden when the package is saved
        [TestMethod]
        public void AutoFilter_of_a_worksheet_writes_the_filters_and_hides_the_rows()
        {
            using (var p = new ExcelPackage())
            {
                var ws = FilterSheet(p);
                Assert.IsNull(ws.AutoFilter, "no autofilter yet");
                ws.Cells["A1:E11"].AutoFilter = true;
                ExcelAutoFilter f = ws.AutoFilter;
                Assert.AreEqual("A1:E11", f.Address.Address);

                // the conditions Excel was given in the reference file; it showed the rows 5 and 8
                f.ShowWhere(4, eFilterOperator.NotEqual, "g0");
                f.ShowValues(0, "alpha", "gamma");
                f.ShowWhere(3, eFilterOperator.Equal, "note*", false, eFilterOperator.Equal, "");
                f.ShowWhere(1, eFilterOperator.GreaterThanOrEqual, 6, true, eFilterOperator.LessThan, 27);
                Assert.IsTrue(f.IsFiltered(0)); Assert.IsFalse(f.IsFiltered(2));
                Assert.AreEqual("", HiddenRows(ws), "nothing is hidden before the save");

                XmlElement sheet = SavedSheet(p);
                XmlElement af = sheet.ChildNodes.OfType<XmlElement>().Single(e => e.LocalName == "autoFilter");
                Assert.AreEqual(
                    "<filterColumn colId=\"0\"><filters><filter val=\"alpha\" /><filter val=\"gamma\" /></filters></filterColumn>" +
                    "<filterColumn colId=\"1\"><customFilters and=\"1\"><customFilter operator=\"greaterThanOrEqual\" val=\"6\" /><customFilter operator=\"lessThan\" val=\"27\" /></customFilters></filterColumn>" +
                    "<filterColumn colId=\"3\"><customFilters><customFilter val=\"note*\" /><customFilter val=\"\" /></customFilters></filterColumn>" +
                    "<filterColumn colId=\"4\"><customFilters><customFilter operator=\"notEqual\" val=\"g0\" /></customFilters></filterColumn>",
                    InnerOf(af));
                Assert.AreEqual("2,3,4,6,7,9,10,11", HiddenRows(ws));

                using (var r = Reload(p))
                {
                    var rs = r.Workbook.Worksheets["S"];
                    Assert.AreEqual("2,3,4,6,7,9,10,11", HiddenRows(rs));
                    Assert.IsTrue(rs.AutoFilter.IsFiltered(4));

                    // a filter taken off after a load: its rows show again
                    rs.AutoFilter.Clear(0);
                    rs.AutoFilter.Clear(1);
                    SavedSheet(r);
                    Assert.AreEqual("4,7,10", HiddenRows(rs), "only the group filter is left");
                    rs.AutoFilter.Clear();
                    rs.AutoFilter.Apply();
                    Assert.AreEqual("", HiddenRows(rs));
                    Assert.AreEqual("", InnerOf(SavedSheet(r).ChildNodes.OfType<XmlElement>().Single(e => e.LocalName == "autoFilter")));
                }
            }
        }

        // #17: the forms Excel uses for empty and not empty, a number as it is shown, and "equal" to a number
        [TestMethod]
        public void AutoFilter_conditions_follow_the_rules_of_Excel()
        {
            // The text a cell shows depends on the culture (12.00 or 12,00), and other tests of this project leave another
            // culture on the thread. The values below are the texts of en-US.
            var culture = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            try
            {
                AutoFilterConditions();
            }
            finally
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = culture;
            }
        }

        static void AutoFilterConditions()
        {
            using (var p = new ExcelPackage())
            {
                // empty cells: rows 5 and 9
                var ws = FilterSheet(p, "A"); ws.Cells["A1:E11"].AutoFilter = true;
                ws.AutoFilter.ShowWhere(3, eFilterOperator.NotEqual, null);
                Assert.AreEqual("<filterColumn colId=\"3\"><customFilters><customFilter operator=\"notEqual\" val=\" \" /></customFilters></filterColumn>",
                    InnerOf(SavedSheet(p, "xl/worksheets/sheet1.xml").ChildNodes.OfType<XmlElement>().Single(e => e.LocalName == "autoFilter")));
                Assert.AreEqual("5,9", HiddenRows(ws), "not empty");
                ws.AutoFilter.ShowValues(3, "note 3", null);
                ws.AutoFilter.Apply();
                Assert.AreEqual("2,3,6,7,8,10,11", HiddenRows(ws), "a value and the empty cells");
                Assert.AreEqual("<filterColumn colId=\"3\"><filters blank=\"1\"><filter val=\"note 3\" /></filters></filterColumn>",
                    InnerOf(SavedSheet(p, "xl/worksheets/sheet1.xml").ChildNodes.OfType<XmlElement>().Single(e => e.LocalName == "autoFilter")));

                // the number 12 is in row 5 and is shown as 12.00
                var b = FilterSheet(p, "B"); b.Cells["A1:E11"].AutoFilter = true;
                b.AutoFilter.ShowValues(1, "12.00", "15.00");
                b.AutoFilter.Apply();
                Assert.AreEqual("2,3,4,7,8,9,10,11", HiddenRows(b), "values are the texts the cells show");
                b.AutoFilter.ShowWhere(1, eFilterOperator.Equal, 12);
                b.AutoFilter.Apply();
                Assert.AreEqual("2,3,4,6,7,8,9,10,11", HiddenRows(b), "equal to a number compares the number");
                Assert.AreEqual("<filterColumn colId=\"1\"><customFilters and=\"1\"><customFilter operator=\"greaterThanOrEqual\" val=\"12\" /><customFilter operator=\"lessThanOrEqual\" val=\"12\" /></customFilters></filterColumn>",
                    InnerOf(SavedSheet(p, "xl/worksheets/sheet2.xml").ChildNodes.OfType<XmlElement>().Single(e => e.LocalName == "autoFilter")));
                b.AutoFilter.ShowWhere(1, eFilterOperator.Equal, 12, false, eFilterOperator.Equal, "15.00");
                b.AutoFilter.Apply();
                Assert.AreEqual("2,3,4,5,7,8,9,10,11", HiddenRows(b), "written as it is, equal is the text the cell shows: 12 is shown as 12.00");
                b.AutoFilter.ShowWhere(1, eFilterOperator.NotEqual, 12);
                b.AutoFilter.Apply();
                Assert.AreEqual("5", HiddenRows(b), "not equal compares the number");

                // dates, wildcards, texts
                b.AutoFilter.Clear();
                b.AutoFilter.ShowWhere(2, eFilterOperator.GreaterThanOrEqual, new DateTime(2026, 1, 6));
                b.AutoFilter.Apply();
                Assert.AreEqual("2,3,4,5,6", HiddenRows(b), "from a date");
                b.AutoFilter.Clear();
                b.AutoFilter.ShowWhere(0, eFilterOperator.Equal, "?l*a");
                b.AutoFilter.Apply();
                Assert.AreEqual("3,4,6,7,9,10,11", HiddenRows(b), "alpha, alpha and Alpha");
                b.AutoFilter.ShowWhere(0, eFilterOperator.GreaterThan, "c");
                b.AutoFilter.Apply();
                Assert.AreEqual("2,3,5,7,8", HiddenRows(b), "after c");
            }
        }

        // #17: the sort is recorded where Excel records it: inside the autofilter of a worksheet, after the autofilter of a table
        [TestMethod]
        public void AutoFilter_sort_moves_the_rows_and_records_the_sort()
        {
            using (var p = new ExcelPackage())
            {
                var ws = FilterSheet(p, "Sheet"); ws.Cells["A1:E11"].AutoFilter = true;
                ws.AutoFilter.ShowWhere(4, eFilterOperator.NotEqual, "g0");
                ws.AutoFilter.Sort(1, true);
                Assert.AreEqual(30, Convert.ToInt32(ws.Cells["B2"].Value)); Assert.AreEqual(3, Convert.ToInt32(ws.Cells["B11"].Value));
                Assert.AreEqual("Name", ws.Cells["A1"].Value, "the row of the buttons stays");
                XmlElement af = SavedSheet(p).ChildNodes.OfType<XmlElement>().Single(e => e.LocalName == "autoFilter");
                Assert.AreEqual("<filterColumn colId=\"4\"><customFilters><customFilter operator=\"notEqual\" val=\"g0\" /></customFilters></filterColumn>" +
                                "<sortState ref=\"A2:E11\"><sortCondition descending=\"1\" ref=\"B2:B11\" /></sortState>", InnerOf(af));
                // the rows of group g0 after the sort: 27, 18, 9 are in the rows 3, 6, 9
                Assert.AreEqual("3,6,9", HiddenRows(ws));
                ws.AutoFilter.Sort(0);
                Assert.AreEqual("<sortCondition ref=\"A2:A11\" />", InnerOf((XmlElement)SavedSheet(p).ChildNodes.OfType<XmlElement>().Single(e => e.LocalName == "autoFilter").LastChild), "one sort, the last one");
                ws.AutoFilter.ClearSort();
                Assert.AreEqual(1, SavedSheet(p).ChildNodes.OfType<XmlElement>().Single(e => e.LocalName == "autoFilter").ChildNodes.Count);

                var ts = FilterSheet(p, "Table");
                var table = ts.Tables.Add(ts.Cells["A1:E11"], "Table_17");
                table.AutoFilter.Sort(1, true);
                table.AutoFilter.ShowValues(0, "alpha", "zeta");
                Assert.AreEqual("zeta", ts.Cells["A2"].Value);
                XmlElement tableXml = SavedSheet(p, "xl/tables/table1.xml");
                string[] names = ChildNames(tableXml);
                CollectionAssert.AreEqual(new[] { "autoFilter", "sortState", "tableColumns", "tableStyleInfo" }, names, string.Join(",", names));
                Assert.AreEqual("<filterColumn colId=\"0\"><filters><filter val=\"alpha\" /><filter val=\"zeta\" /></filters></filterColumn>", InnerOf((XmlElement)tableXml.FirstChild));
                Assert.AreEqual("<sortCondition descending=\"1\" ref=\"B2:B11\" />", InnerOf((XmlElement)tableXml.ChildNodes[1]));
                Assert.AreEqual("A2:E11", ((XmlElement)tableXml.ChildNodes[1]).GetAttribute("ref"));
                // sorted by Qty, largest first: zeta 30, gamma 27, epsilon 24, Alpha 21, beta 18, delta 15, alpha 12, gamma 9, beta 6, alpha 3
                Assert.AreEqual("3,4,6,7,9,10", HiddenRows(ts));

                using (var r = Reload(p))
                {
                    Assert.AreEqual("3,4,6,7,9,10", HiddenRows(r.Workbook.Worksheets["Table"]));
                    Assert.IsTrue(r.Workbook.Worksheets["Table"].Tables[0].AutoFilter.IsFiltered(0));
                }
            }
        }

        // #17: a filter that is not made here is left to Excel; nothing changes on a sheet where the filters were not touched
        [TestMethod]
        public void AutoFilter_leaves_other_filters_and_untouched_sheets_alone()
        {
            using (var p = new ExcelPackage())
            {
                var ws = FilterSheet(p); ws.Cells["A1:E11"].AutoFilter = true;
                // as a file from Excel with a top 3 filter on Qty: the rows 2 to 8 are hidden
                XmlElement af = (XmlElement)ws.WorksheetXml.DocumentElement.SelectSingleNode("d:autoFilter", ws.NameSpaceManager);
                af.InnerXml = "<filterColumn colId=\"1\" xmlns=\"" + ExcelPackage.schemaMain + "\"><top10 val=\"3\" filterVal=\"24\" /></filterColumn>";
                for (int r = 2; r <= 8; r++) ws.Row(r).Hidden = true;

                var untouched = FilterSheet(p, "Untouched"); untouched.Cells["A1:E11"].AutoFilter = true;
                untouched.Row(4).Hidden = true;
                var plain = FilterSheet(p, "Plain");
                var table = plain.Tables.Add(plain.Cells["A1:E11"], "Table_Plain");
                table.ShowFilter = false;
                Assert.IsNull(table.AutoFilter, "no filter buttons");

                using (var r = Reload(p))
                {
                    var rs = r.Workbook.Worksheets["S"];
                    Assert.AreEqual("2,3,4,5,6,7,8", HiddenRows(rs));
                    Assert.AreEqual("4", HiddenRows(r.Workbook.Worksheets["Untouched"]), "a row hidden by hand stays hidden");

                    // one more filter on another column: the rows of the top 3 filter stay hidden, the new filter hides more
                    rs.AutoFilter.ShowWhere(4, eFilterOperator.NotEqual, "g0");
                    XmlElement saved = SavedSheet(r).ChildNodes.OfType<XmlElement>().Single(e => e.LocalName == "autoFilter");
                    Assert.AreEqual("<filterColumn colId=\"1\"><top10 val=\"3\" filterVal=\"24\" /></filterColumn>" +
                                    "<filterColumn colId=\"4\"><customFilters><customFilter operator=\"notEqual\" val=\"g0\" /></customFilters></filterColumn>", InnerOf(saved));
                    Assert.AreEqual("2,3,4,5,6,7,8,10", HiddenRows(rs));
                    Assert.AreEqual("4", HiddenRows(r.Workbook.Worksheets["Untouched"]));
                }
            }
        }

        // #17: the totals row of a table is below the autofilter: it is not sorted with the rows and not hidden
        [TestMethod]
        public void AutoFilter_of_a_table_leaves_the_totals_row_alone()
        {
            using (var p = new ExcelPackage())
            {
                var ws = FilterSheet(p);
                var table = ws.Tables.Add(ws.Cells["A1:E11"], "Table_Totals");
                table.ShowTotal = true;
                ws.Cells["A12"].Value = "Total";
                Assert.AreEqual("A1:E12", table.Address.Address);
                Assert.AreEqual("A1:E11", table.AutoFilter.Address.Address);

                table.AutoFilter.Sort(0, true);
                table.AutoFilter.ShowValues(0, "nothing has this name");
                table.AutoFilter.Apply();
                Assert.AreEqual("Total", ws.Cells["A12"].Value, "not sorted");
                Assert.AreEqual("zeta", ws.Cells["A2"].Value);
                Assert.AreEqual("2,3,4,5,6,7,8,9,10,11", HiddenRows(ws, 2, 12), "every row but the totals row");
            }
        }

        // #17: wrong arguments, and an autofilter that was removed
        [TestMethod]
        public void AutoFilter_refuses_wrong_arguments()
        {
            using (var p = new ExcelPackage())
            {
                var ws = FilterSheet(p); ws.Cells["A1:E11"].AutoFilter = true;
                ExcelAutoFilter f = ws.AutoFilter;
                Assert.ThrowsException<ArgumentOutOfRangeException>(() => f.ShowValues(5, "x"));
                Assert.ThrowsException<ArgumentOutOfRangeException>(() => f.ShowWhere(-1, eFilterOperator.Equal, 1));
                Assert.ThrowsException<ArgumentOutOfRangeException>(() => f.Sort(9));
                Assert.ThrowsException<ArgumentException>(() => f.ShowValues(0));
                f.ShowValues(0, "alpha");
                ws.Cells["A1:E11"].AutoFilter = false;
                Assert.IsNull(ws.AutoFilter);
                Assert.ThrowsException<InvalidOperationException>(() => f.ShowValues(0, "beta"));
                // saved without the autofilter: the filter that was set on it does nothing
                Assert.IsFalse(ChildNames(SavedSheet(p)).Contains("autoFilter"));
                Assert.AreEqual("", HiddenRows(ws));
            }
        }

        // #14: AutoFitColumns keeps what it measured and counts digits as one character when the font allows it.
        // The width of a column must still be exactly the width of its widest cell measured on its own.
        [TestMethod]
        public void AutoFit_of_a_range_gives_each_column_its_widest_cell_measured_alone()
        {
            // Calibri has digits of one width; Georgia has digits of several widths, so its digits can not count as one
            foreach (string font in new[] { "Calibri", "Georgia", "Consolas" })
            {
                using (var p = new ExcelPackage())
                {
                    var ws = p.Workbook.Worksheets.Add("S");
                    const int rows = 1300, cols = 7;
                    var rnd = new Random(14);
                    for (int r = 1; r <= rows; r++)
                    {
                        ws.Cells[r, 1].Value = rnd.Next(100000000);
                        ws.Cells[r, 2].Value = Math.Round(rnd.NextDouble() * 1000000, 3);
                        ws.Cells[r, 3].Value = new DateTime(2020, 1, 1).AddSeconds(rnd.Next(200000000));
                        ws.Cells[r, 4].Value = "id " + rnd.Next(500) + "." + rnd.Next(10) + ":" + rnd.Next(60) + "-" + rnd.Next(9) + "," + rnd.Next(99) + " (" + rnd.Next(7) + "/" + rnd.Next(7) + ")";
                        ws.Cells[r, 5].Value = "세션 " + rnd.Next(100000) + " 대기";
                        ws.Cells[r, 6].Value = "SQL_" + rnd.Next().ToString("X8");
                        ws.Cells[r, 7].Value = rnd.Next(3) == 0 ? "17" + rnd.Next(100) : "name " + rnd.Next(40);
                    }
                    ws.Cells[1, 1, rows, cols].Style.Font.Name = font;
                    ws.Cells[1, 2, rows, 2].Style.Numberformat.Format = "#,##0.000";
                    ws.Cells[1, 3, rows, 3].Style.Numberformat.Format = "yyyy-mm-dd hh:mm:ss";
                    ws.Cells[1, 7, rows, 7].Style.Font.Bold = true;

                    // one row at a time: so few texts that nothing is taken from another cell
                    var expected = new double[cols];
                    for (int r = 1; r <= rows; r++)
                    {
                        ws.Cells[r, 1, r, cols].AutoFitColumns(0);
                        for (int c = 0; c < cols; c++) expected[c] = Math.Max(expected[c], ws.Column(c + 1).Width);
                    }

                    ws.Cells[1, 1, rows, cols].AutoFitColumns(0);
                    for (int c = 0; c < cols; c++)
                    {
                        Assert.IsTrue(expected[c] > 1, font + " column " + (c + 1));
                        Assert.AreEqual(expected[c], ws.Column(c + 1).Width, font + " column " + (c + 1));
                    }
                }
            }
        }

        // #14: the rules of AutoFitColumns that the faster loop has to keep
        [TestMethod]
        public void AutoFit_keeps_its_rules_for_hidden_wrapped_merged_filter_and_limit()
        {
            const string text = "a fairly long text that sets the width of its column";
            using (var p = new ExcelPackage())
            {
                var plain = p.Workbook.Worksheets.Add("Plain");
                plain.Cells["A1"].Value = text;
                plain.Cells["A1"].AutoFitColumns(0);
                double width = plain.Column(1).Width;
                Assert.IsTrue(width > 30, width.ToString());

                var ws = p.Workbook.Worksheets.Add("S");
                for (int c = 1; c <= 7; c++) ws.Cells[1, c].Value = text;
                ws.Cells["A2"].Value = "x";
                ws.Column(2).Width = 3; ws.Column(2).Hidden = true;      // B: hidden
                ws.Cells["C1"].Style.WrapText = true;                     // C: wrapped
                ws.Cells["D1:E1"].Merge = true;                           // D, E: merged
                ws.Cells["F1:F3"].AutoFilter = true;                      // F: filter button
                ws.Cells["A1:F3"].AutoFitColumns(0);
                Assert.AreEqual(width, ws.Column(1).Width, "plain");
                Assert.IsTrue(ws.Column(2).Hidden, "hidden stays hidden");
                Assert.AreEqual(0d, ws.Column(2).Width, "hidden gets the minimum width and is not widened");
                Assert.IsTrue(ws.Column(3).Width < width / 2, "wrapped: " + ws.Column(3).Width);
                Assert.IsTrue(ws.Column(4).Width < width / 2, "merged: " + ws.Column(4).Width);
                Assert.AreEqual(width + 2.25, ws.Column(6).Width, "filter button");
                Assert.AreEqual(9.140625, ws.Column(7).Width, 0.5, "outside the range");

                ws.Cells["A1:A2"].AutoFitColumns(2, 12);
                Assert.AreEqual(12d, ws.Column(1).Width, "limit");

                // a range of several addresses, and a column that has no cell in it
                var multi = p.Workbook.Worksheets.Add("Multi");
                multi.Cells["A1"].Value = text; multi.Cells["C1"].Value = "short"; multi.Cells["E1"].Value = text;
                multi.Cells["A1:A5,E1:E5"].AutoFitColumns(0);
                Assert.AreEqual(width, multi.Column(1).Width, "first address");
                Assert.AreEqual(width, multi.Column(5).Width, "second address");
            }
        }

        // #9: the element is placed by the elements before it, so later elements that the library does not know do not matter
        [TestMethod]
        public void Ignored_errors_come_before_later_elements_the_library_does_not_know()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("S");
                ws.Cells["A1"].Value = "001";
                XmlElement top = ws.WorksheetXml.DocumentElement;
                // sheetCalcPr comes right after sheetData and before ignoredErrors; controls comes after ignoredErrors. Neither is in the node order list of the worksheet.
                XmlNode sheetData = top.SelectSingleNode("d:sheetData", ws.NameSpaceManager);
                top.InsertAfter(top.OwnerDocument.CreateElement("sheetCalcPr", ExcelPackage.schemaMain), sheetData);
                top.AppendChild(top.OwnerDocument.CreateElement("controls", ExcelPackage.schemaMain));

                ws.IgnoredErrors.Add("A1", eErrorChecks.NumberStoredAsText);
                string[] names = ChildNames(SavedSheet(p));
                int ix = Array.IndexOf(names, "ignoredErrors");
                Assert.IsTrue(ix > Array.IndexOf(names, "sheetCalcPr") && ix > Array.IndexOf(names, "pageMargins"), string.Join(",", names));
                Assert.IsTrue(ix < Array.IndexOf(names, "controls"), string.Join(",", names));
            }
        }

        // #9: a range object with several ranges, a range of another sheet, and taking every check away
        [TestMethod]
        public void Ignored_errors_take_range_objects_of_their_own_sheet()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("S");
                var other = p.Workbook.Worksheets.Add("Other");
                ws.Cells["A1"].Value = 1; other.Cells["A1"].Value = 1;

                var a = ws.IgnoredErrors.Add(ws.Cells["A1:A3,C1:C3"], eErrorChecks.NumberStoredAsText);
                Assert.AreEqual("A1:A3 C1:C3", a.Address);
                var b = ws.IgnoredErrors.Add(new ExcelAddress("E:E"), eErrorChecks.EvalError);
                Assert.AreEqual("E1:E1048576", b.Address);
                Assert.ThrowsException<ArgumentException>(() => ws.IgnoredErrors.Add(other.Cells["A1"], eErrorChecks.All));
                Assert.AreEqual(2, ws.IgnoredErrors.Count);

                Assert.ThrowsException<ArgumentException>(() => a.Checks = eErrorChecks.None);
                Assert.AreEqual(eErrorChecks.NumberStoredAsText, a.Checks);
            }
        }

        // #9: an entry of a loaded file whose address cannot be read is left alone; inserting and deleting still work
        [TestMethod]
        public void Ignored_errors_with_an_unreadable_address_do_not_stop_an_insert()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("S");
                ws.Cells["A1"].Value = 1; ws.Cells["A9"].Value = 2;
                ws.IgnoredErrors.Add("A5:A9", eErrorChecks.NumberStoredAsText);
                var odd = ws.IgnoredErrors.Add("B1", eErrorChecks.EvalError);
                odd.Node.SetAttribute("sqref", "NotARange");
                using (var r = Reload(p))
                {
                    var rs = r.Workbook.Worksheets["S"];
                    rs.InsertRow(2, 3);
                    rs.DeleteColumn(4, 1);
                    Assert.AreEqual("A8:A12", rs.IgnoredErrors[0].Address);
                    Assert.AreEqual("NotARange", rs.IgnoredErrors[1].Address);
                    Assert.AreEqual(eErrorChecks.EvalError, rs.IgnoredErrors[1].Checks);
                }
            }
        }
    }
}
