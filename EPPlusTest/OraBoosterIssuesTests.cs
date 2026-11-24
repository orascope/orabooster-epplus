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
