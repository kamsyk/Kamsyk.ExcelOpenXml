using DocumentFormat.OpenXml;
//using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Kamsyk.ExcelOpenXml {
    /// <summary>
    /// Generate xlsx files
    /// </summary>
    public class Excel {
        #region Constants
        public static DateTime ZERO_DATE = new DateTime(1899, 12, 30);
        #endregion

        #region Enumerators
        public enum ChartType {
            xlLine = 4,
            xlColumnClustered = 51,
            xlPie = 5
        }
        #endregion

        #region Properties
        private MemoryStream m_spreadsheetStream;
        #endregion

        #region Methods
        public SpreadsheetDocument GetNewXlsDocMemory() {
            var spreadsheetStream = new MemoryStream();
            m_spreadsheetStream = new MemoryStream();
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(m_spreadsheetStream, SpreadsheetDocumentType.Workbook);


            return spreadsheetDocument;
        }

        public WorkbookPart AddWorkbook(SpreadsheetDocument spreadsheetDocument) {

            // Add a WorksheetPart to the WorkbookPart.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            SharedStringTablePart shareStringPart = workbookpart.AddNewPart<SharedStringTablePart>();
            shareStringPart.SharedStringTable = new SharedStringTable();

            return workbookpart;
        }

        public WorksheetPart AddWorkSheet(WorkbookPart workbookpart) {
            return AddWorkSheet(workbookpart, null);
        }

        public WorksheetPart AddWorkSheet(WorkbookPart workbookpart, string sheetName) {
            // Append a new worksheet and associate it with the workbook.
            //int sheetId = 1;
            //if (workbookpart.Workbook.Sheets != null) {
            //    sheetId = workbookpart.Workbook.Sheets.Count() + 1;
            //}
            //uint uiSheetId = Convert.ToUInt32(sheetId);
            //if (String.IsNullOrEmpty(sheetName)) {
            //    sheetName = "Sheet " + sheetId;
            //}

            //Sheet sheet = new Sheet() {
            //    Id = workbookpart.GetIdOfPart(workbookpart.WorksheetParts.LastOrDefault()), SheetId = uiSheetId, Name = sheetName
            //};
            //workbookpart.Workbook.Sheets.Append(sheet);

            //return workbookpart.WorksheetParts.LastOrDefault();

            // Add a blank WorksheetPart.
            WorksheetPart newWorksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());


            Sheets sheets = workbookpart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookpart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new worksheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0) {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            //// Give the new worksheet a name.
            //string sheetName = "Sheet" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            Sheet newSheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(newSheet);

            return newWorksheetPart;
        }

        public WorksheetPart AddWorkSheet(WorkbookPart workbookpart, string sheetName, int wsIndex) {
            // Add a blank WorksheetPart.
            WorksheetPart newWorksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());


            Sheets sheets = workbookpart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookpart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new worksheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0) {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            var sheet = workbookpart.Workbook.Sheets.ElementAt(wsIndex);

            // Append the new worksheet and associate it with the workbook.
            Sheet newSheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.InsertBefore(newSheet, sheet);

            return newWorksheetPart;
        }

        public MemoryStream GetWorkbookMemoryStream() {
            m_spreadsheetStream.Position = 0;
            return m_spreadsheetStream;
        }

        public MemoryStream GenerateExcelWorkbook(System.Data.DataTable dataTable) {
            return GenerateExcelWorkbook(dataTable, null);
        }
        public MemoryStream GenerateExcelWorkbook(System.Data.DataTable dataTable, List<double> columnWidths) {
            SpreadsheetDocument xlsDoc = GenerateExcelWorkbookDoc(dataTable, columnWidths);

            xlsDoc.Close();

            return GetWorkbookMemoryStream();
        }

        public SpreadsheetDocument GenerateExcelWorkbookDoc(System.Data.DataTable dataTable, List<double> columnWidths) {

            var xlsDoc = GetNewXlsDocMemory();
            var wbPart = AddWorkbook(xlsDoc);
            string sheetName = (String.IsNullOrEmpty(dataTable.TableName)) ? "Sheet1" : dataTable.TableName;
            var wsPart = AddWorkSheet(wbPart, sheetName);

            var worksheetParts = wbPart.WorksheetParts;
            var ws = wsPart.Worksheet;

            if (columnWidths != null) {
                Columns columns = new Columns();

                for (int i = 0; i < columnWidths.Count; i++) {
                    uint iIndex = Convert.ToUInt32(i + 1);
                    columns.Append(CreateColumnData(iIndex, iIndex, columnWidths[i]));
                }

                var sheetdata = wsPart.Worksheet.GetFirstChild<SheetData>();
                wsPart.Worksheet.InsertBefore(columns, sheetdata);
            }


            int iRow = 1;

            var stylesPart = xlsDoc.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet();

            // blank font list
            stylesPart.Stylesheet.Fonts = new DocumentFormat.OpenXml.Spreadsheet.Fonts();
            stylesPart.Stylesheet.Fonts.Count = 2;
            stylesPart.Stylesheet.Fonts.AppendChild(new Font());
            Font headeFont = new Font(new Bold());
            //headeFont.Bold = new Bold();
            stylesPart.Stylesheet.Fonts.AppendChild(headeFont);

            // create fills
            stylesPart.Stylesheet.Fills = new Fills();


            // create a solid red fill
            var solidGreen = new DocumentFormat.OpenXml.Spreadsheet.PatternFill() { PatternType = PatternValues.Solid };
            solidGreen.ForegroundColor = new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor { Rgb = HexBinaryValue.FromString("66ff33") }; // red fill
            solidGreen.BackgroundColor = new DocumentFormat.OpenXml.Spreadsheet.BackgroundColor { Indexed = 64 };

            stylesPart.Stylesheet.Fills.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Fill { PatternFill = new DocumentFormat.OpenXml.Spreadsheet.PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
            stylesPart.Stylesheet.Fills.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Fill { PatternFill = new DocumentFormat.OpenXml.Spreadsheet.PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
            stylesPart.Stylesheet.Fills.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Fill { PatternFill = solidGreen });
            stylesPart.Stylesheet.Fills.Count = 3;

            // blank border list
            stylesPart.Stylesheet.Borders = new Borders();
            stylesPart.Stylesheet.Borders.Count = 1;
            stylesPart.Stylesheet.Borders.AppendChild(new Border());

            // blank cell format list
            stylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();
            stylesPart.Stylesheet.CellStyleFormats.Count = 1;
            stylesPart.Stylesheet.CellStyleFormats.AppendChild(new CellFormat());

            ////number formats
            //stylesPart.Stylesheet.NumberingFormats = new NumberingFormats();
            //stylesPart.Stylesheet.NumberingFormats.Count = 1;
            //stylesPart.Stylesheet.NumberingFormats.AppendChild(new CellFormat());


            // cell format list
            stylesPart.Stylesheet.CellFormats = new CellFormats();
            // empty one for index 0, seems to be required
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat());
            // cell format references style format 0, font 0, border 0, fill 2 and applies the fill
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 1, BorderId = 0, FillId = 2, ApplyFill = true }).AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Left });
            stylesPart.Stylesheet.CellFormats.Count = 2;

            stylesPart.Stylesheet.Save();
            SheetData sheetData = ws.GetFirstChild<SheetData>();

            for (int i = 0; i < dataTable.Columns.Count; i++) {
                string cellAddress = GetColumnAddress(i + 1) + iRow;
                UInt32 rowNumber = Convert.ToUInt32(iRow);
                Row row = GetRow(sheetData, rowNumber);
                var cell = InsertNewCellInWorksheet(ws, row, cellAddress);
                SetCellValue(cell, dataTable.Columns[i].ColumnName);
                cell.StyleIndex = 1;

            }
            iRow++;

            //Data
            for (int i = 0; i < dataTable.Rows.Count; i++) {

                UInt32 rowNumber = Convert.ToUInt32(iRow);
                Row row = GetRow(sheetData, rowNumber);

                for (int j = 0; j < dataTable.Columns.Count; j++) {
                    if (dataTable.Rows[i][j] == null || dataTable.Rows[i][j] == DBNull.Value ||
                        dataTable.Rows[i][j].ToString().Trim().Length == 0) {
                        continue;
                    }
                    string cellAddress = GetColumnAddress(j + 1) + iRow;
                    var cell = InsertNewCellInWorksheet(ws, row, cellAddress);
                    SetCellValue(cell, dataTable.Rows[i][j]);
                    //int unicode = 6;
                    //char character = (char)unicode;
                    //string text = character.ToString();
                    //if (dataTable.Rows[i][j].ToString().Contains(text)) {
                    //    int sdaasd = 5;
                    //}
                }

                iRow++;
            }


            //xlsDoc.Close();

            return xlsDoc;
        }

        public MemoryStream GenerateExcelWorkbook(System.Data.DataTable dataTable, ChartType chartType, string chartTitle) {
            if (chartType == ChartType.xlLine) {
                return GenerateExcelChartLine(dataTable, chartTitle);
            }

            if (chartType == ChartType.xlPie) {
                return GenerateExcelChartPie(dataTable, chartTitle);
            }

            return GenerateExcelChartBar(dataTable, chartTitle);
        }

        private MemoryStream GenerateExcelChartBar(System.Data.DataTable dataTable, string chartTitle) {
            List<double> columnWidths = new List<double>();
            columnWidths.Add(50);
            for (int i = 1; i < dataTable.Columns.Count; i++) {
                columnWidths.Add(30);
            }

            SpreadsheetDocument xlsDoc = GenerateExcelWorkbookDoc(dataTable, columnWidths);

            string sheetName = "Chart";
            var wbPart = xlsDoc.WorkbookPart;
            var wsPart = AddWorkSheet(wbPart, sheetName, 0);

            #region Charts
            // Add a new drawing to the worksheet
            DrawingsPart drawingsPart = wsPart.AddNewPart<DrawingsPart>();
            wsPart.Worksheet.Append(new Drawing() { Id = wsPart.GetIdOfPart(drawingsPart) });
            wsPart.Worksheet.Save();

            drawingsPart.WorksheetDrawing = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
            // Add a new chart and set the chart language
            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new ChartSpace();
            chartPart.ChartSpace.AppendChild(new EditingLanguage() { Val = "en-US" });
            Chart chart = chartPart.ChartSpace.AppendChild(new Chart());
            
            //chart.AppendChild(new AutoTitleDeleted() { Val = true }); // We don't want to show the chart title

            // Create a new Clustered Column Chart
            PlotArea plotArea = chart.AppendChild(new PlotArea());
            Layout layout = plotArea.AppendChild(new Layout());

            BarChart barChart = plotArea.AppendChild(new BarChart(
                    new BarDirection() { Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column) },
                    new BarGrouping() { Val = new EnumValue<BarGroupingValues>(BarGroupingValues.Clustered) },
                    new VaryColors() { Val = false }
                ));
            
            // Create chart series
            string lastCol = CommonExcel.GetLetterFromColIndex(dataTable.Columns.Count);
            for (int i = 0; i < dataTable.Rows.Count; i++) {
                BarChartSeries barChartSeries = barChart.AppendChild(new BarChartSeries(
                    new Index() { Val = (uint)i },
                    new Order() { Val = (uint)i },
                    new SeriesText(new NumericValue() { Text = dataTable.Rows[i][0].ToString() })
                ));

                // Adding category axis to the chart
                CategoryAxisData categoryAxisData = barChartSeries.AppendChild(new CategoryAxisData());

                // Category
                // Constructing the chart category
                //string lastCol = CommonExcel.GetLetterFromColIndex(dataTable.Columns.Count - 1);
                string formulaCat = "Data!$B$1:$" + lastCol + "$1";

                StringReference stringReference = categoryAxisData.AppendChild(new StringReference() {
                    Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formulaCat }
                });

                StringCache stringCache = stringReference.AppendChild(new StringCache());
                stringCache.Append(new PointCount() { Val = (uint)(dataTable.Columns.Count - 1) });

                for (int j = 1; j < (dataTable.Columns.Count); j++) {
                    stringCache.AppendChild(new NumericPoint() { Index = (uint)j }).Append(new NumericValue(dataTable.Columns[j].ColumnName));
                }
            }

            //var chartSeries = barChart.Elements<BarChartSeries>().GetEnumerator();
            var chartSeries = barChart.Elements<BarChartSeries>();
            
            for (int i = 0; i < dataTable.Rows.Count; i++) {
                BarChartSeries bcs = chartSeries.ElementAt(i);

                string formulaVal = string.Format("Data!$B${0}:$" + lastCol + "${0}", (i + 2));
                DocumentFormat.OpenXml.Drawing.Charts.Values values = bcs.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Values());

                NumberReference numberReference = values.AppendChild(new NumberReference() {
                    Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formulaVal }
                });

                NumberingCache numberingCache = numberReference.AppendChild(new NumberingCache());
                numberingCache.Append(new PointCount() { Val = (uint)(dataTable.Columns.Count -1) });

                for (int j = 1; j < dataTable.Columns.Count; j++) {
                    numberingCache.AppendChild(new NumericPoint() { Index = (uint)j }).Append(new NumericValue(dataTable.Rows[i][j].ToString()));
                }

                
            }

            barChart.AppendChild(new DataLabels(
                                new ShowLegendKey() { Val = true },
                                new ShowValue() { Val = true },
                                new ShowCategoryName() { Val = false },
                                new ShowSeriesName() { Val = false },
                                new ShowPercent() { Val = false },
                                new ShowBubbleSize() { Val = false }
                            ));

            barChart.Append(new AxisId() { Val = 48650112u });
            barChart.Append(new AxisId() { Val = 48672768u });

            // Adding Category Axis
            plotArea.AppendChild(
                new CategoryAxis(
                    new AxisId() { Val = 48650112u },
                    new Scaling(new Orientation() { Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax) }),
                    new Delete() { Val = false },
                    new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
                    new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                    new CrossingAxis() { Val = 48672768u },
                    new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                    new AutoLabeled() { Val = true },
                    new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) }
                ));

            // Adding Value Axis
            plotArea.AppendChild(
                new ValueAxis(
                    new AxisId() { Val = 48672768u },
                    new Scaling(new Orientation() { Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax) }),
                    new Delete() { Val = false },
                    new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
                    new MajorGridlines(),
                    new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat() {
                        FormatCode = "General",
                        SourceLinked = true
                    },
                    new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                    new CrossingAxis() { Val = 48650112u },
                    new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                    new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }
                ));

            chart.Append(
                    new PlotVisibleOnly() { Val = true },
                    new DisplayBlanksAs() { Val = new EnumValue<DisplayBlanksAsValues>(DisplayBlanksAsValues.Gap) },
                    new ShowDataLabelsOverMaximum() { Val = false }
                );

            chartPart.ChartSpace.Save();

            // Positioning the chart on the spreadsheet
            TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild(new TwoCellAnchor());

            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(
                    new ColumnId("0"),
                    new ColumnOffset("0"),
                    new RowId((1).ToString()),
                    new RowOffset("0")
                ));

            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(
                    new ColumnId("20"),
                    new ColumnOffset("0"),
                    new RowId((35).ToString()),
                    new RowOffset("0")
                ));

            // Append GraphicFrame to TwoCellAnchor
            GraphicFrame graphicFrame = twoCellAnchor.AppendChild(new GraphicFrame());
            graphicFrame.Macro = string.Empty;

            graphicFrame.Append(new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties() {
                        Id = 2u,
                        Name = "Reget Statistics"
                    },
                    new NonVisualGraphicFrameDrawingProperties()
                ));

            graphicFrame.Append(new Transform(
                    new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                    new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 0L, Cy = 0L }
                ));

            graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Graphic(
                    new DocumentFormat.OpenXml.Drawing.GraphicData(
                            new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) }
                        ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                ));

            twoCellAnchor.Append(new ClientData());

            AddChartTitle(chart, chartTitle);
            #endregion

            #region Chart
            //    DrawingsPart drawingsPart = wsPart.AddNewPart<DrawingsPart>();
            //    wsPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing() { Id = wsPart.GetIdOfPart(drawingsPart) });
            //    wsPart.Worksheet.Save();

            //    ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
            //    chartPart.ChartSpace = new ChartSpace();
            //    chartPart.ChartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });
            //    DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartPart.ChartSpace.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>(
            //        new DocumentFormat.OpenXml.Drawing.Charts.Chart());

            //    // Create a new clustered column chart.
            //    PlotArea plotArea = chart.AppendChild<PlotArea>(new PlotArea());
            //    Layout layout = plotArea.AppendChild<Layout>(new Layout());
            //    BarChart barChart = plotArea.AppendChild<BarChart>(new BarChart(new BarDirection() { Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column) },
            //        new BarGrouping() { Val = new EnumValue<BarGroupingValues>(BarGroupingValues.Clustered) }));

            //    //uint i = 0;
            //    for (int iRow = 0; iRow < dataTable.Rows.Count; iRow++) {
            //        uint i = (uint)iRow;
            //        BarChartSeries barChartSeries = barChart.AppendChild<BarChartSeries>(
            //            new BarChartSeries(new Index() { Val = new UInt32Value(i)},
            //            new Order() { Val = new UInt32Value(i) },
            //            new SeriesText(new NumericValue() { Text = dataTable.Rows[iRow][0].ToString() })));

            //        StringLiteral strLit = barChartSeries.AppendChild<CategoryAxisData>(new CategoryAxisData()).AppendChild<StringLiteral>(new StringLiteral());
            //        strLit.Append(new PointCount() { Val = new UInt32Value(1U) });
            //        strLit.AppendChild<StringPoint>(new StringPoint() { Index = new UInt32Value(0U) }).Append(new NumericValue("title" + iRow));

            //        NumberLiteral numLit = barChartSeries.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Values>(
            //            new DocumentFormat.OpenXml.Drawing.Charts.Values()).AppendChild<NumberLiteral>(new NumberLiteral());
            //        numLit.Append(new FormatCode("General"));
            //        numLit.Append(new PointCount() { Val = new UInt32Value(1U) });
            //        numLit.AppendChild<NumericPoint>(new NumericPoint() { Index = new UInt32Value(0u) }).Append(new NumericValue(dataTable.Rows[iRow][1].ToString()));

            //        barChart.Append(new AxisId() { Val = new UInt32Value(48650112u) });
            //        barChart.Append(new AxisId() { Val = new UInt32Value(48672768u) });

            //        // Add the Category Axis.
            //        CategoryAxis catAx = plotArea.AppendChild<CategoryAxis>(new CategoryAxis(new AxisId() { Val = new UInt32Value(48650112u) }, new Scaling(new Orientation() {
            //            Val = new EnumValue<DocumentFormat.
            //OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
            //        }),
            //            new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
            //            new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
            //            new CrossingAxis() { Val = new UInt32Value(48672768U) },
            //            new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
            //            new AutoLabeled() { Val = new BooleanValue(true) },
            //            new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) },
            //            new LabelOffset() { Val = new UInt16Value((ushort)100) }));

            //        // Add the Value Axis.
            //        ValueAxis valAx = plotArea.AppendChild<ValueAxis>(new ValueAxis(new AxisId() { Val = new UInt32Value(48672768u) },
            //            new Scaling(new Orientation() {
            //                Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(
            //                DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
            //            }),
            //            new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
            //            new MajorGridlines(),
            //            new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat() {
            //                FormatCode = new StringValue("General"),
            //                SourceLinked = new BooleanValue(true)
            //            }, new TickLabelPosition() {
            //                Val = new EnumValue<TickLabelPositionValues>
            //(TickLabelPositionValues.NextTo)
            //            }, new CrossingAxis() { Val = new UInt32Value(48650112U) },
            //            new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
            //            new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }));

            //        // Add the chart Legend.
            //        Legend legend = chart.AppendChild<Legend>(new Legend(new LegendPosition() { Val = new EnumValue<LegendPositionValues>(LegendPositionValues.Right) },
            //            new Layout()));

            //        chart.Append(new PlotVisibleOnly() { Val = new BooleanValue(true) });

            //        // Save the chart part.
            //        chartPart.ChartSpace.Save();

            //        // Position the chart on the worksheet using a TwoCellAnchor object.
            //        drawingsPart.WorksheetDrawing = new WorksheetDrawing();
            //        TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild<TwoCellAnchor>(new TwoCellAnchor());
            //        twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(new ColumnId("9"),
            //            new ColumnOffset("581025"),
            //            new RowId("17"),
            //            new RowOffset("114300")));
            //        twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(new ColumnId("17"),
            //            new ColumnOffset("276225"),
            //            new RowId("32"),
            //            new RowOffset("0")));

            //        // Append a GraphicFrame to the TwoCellAnchor object.
            //        DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame = twoCellAnchor.AppendChild<DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame>(new DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame());
            //        graphicFrame.Macro = "";

            //        graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
            //            new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties() { Id = new UInt32Value(2u), Name = "Chart 1" },
            //            new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()));

            //        graphicFrame.Append(new Transform(new Offset() { X = 0L, Y = 0L },
            //                                                                new Extents() { Cx = 0L, Cy = 0L }));

            //        graphicFrame.Append(new Graphic(new GraphicData(new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) }) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }));

            //        twoCellAnchor.Append(new ClientData());

            //    }
            #endregion

            xlsDoc.Close();

            return GetWorkbookMemoryStream();
        }

        private MemoryStream GenerateExcelChartLine(System.Data.DataTable dataTable, string chartTitle) {
            List<double> columnWidths = new List<double>();
            columnWidths.Add(50);
            for (int i = 1; i < dataTable.Columns.Count; i++) {
                columnWidths.Add(30);
            }

            SpreadsheetDocument xlsDoc = GenerateExcelWorkbookDoc(dataTable, columnWidths);

            string sheetName = "Chart";
            var wbPart = xlsDoc.WorkbookPart;
            var wsPart = AddWorkSheet(wbPart, sheetName, 0);

            #region Charts
            // Add a new drawing to the worksheet
            DrawingsPart drawingsPart = wsPart.AddNewPart<DrawingsPart>();
            wsPart.Worksheet.Append(new Drawing() { Id = wsPart.GetIdOfPart(drawingsPart) });
            wsPart.Worksheet.Save();

            drawingsPart.WorksheetDrawing = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
            // Add a new chart and set the chart language
            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new ChartSpace();
            chartPart.ChartSpace.AppendChild(new EditingLanguage() { Val = "en-US" });
            Chart chart = chartPart.ChartSpace.AppendChild(new Chart());

            
            // Create a new Clustered Column Chart
            PlotArea plotArea = chart.AppendChild(new PlotArea());
            Layout layout = plotArea.AppendChild(new Layout());

            LineChart lineChart = plotArea.AppendChild(new LineChart());

            // Create chart series
            string lastCol = CommonExcel.GetLetterFromColIndex(dataTable.Columns.Count);
            for (int i = 0; i < dataTable.Rows.Count; i++) {
                LineChartSeries lineChartSeries = lineChart.AppendChild(new LineChartSeries(
                    new Index() { Val = (uint)i },
                    new Order() { Val = (uint)i },
                    new SeriesText(new NumericValue() { Text = dataTable.Rows[i][0].ToString() })
                ));

                // Adding category axis to the chart
                CategoryAxisData categoryAxisData = lineChartSeries.AppendChild(new CategoryAxisData());

                // Category
                // Constructing the chart category
                string formulaCat = "Data!$B$1:$" + lastCol + "$1";

                StringReference stringReference = categoryAxisData.AppendChild(new StringReference() {
                    Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formulaCat }
                });

                StringCache stringCache = stringReference.AppendChild(new StringCache());
                stringCache.Append(new PointCount() { Val = (uint)(dataTable.Columns.Count - 1) });

                for (int j = 1; j < (dataTable.Columns.Count); j++) {
                    stringCache.AppendChild(new NumericPoint() { Index = (uint)j }).Append(new NumericValue(dataTable.Columns[j].ColumnName));
                }
            }

            var chartSeries = lineChart.Elements<LineChartSeries>();

            for (int i = 0; i < dataTable.Rows.Count; i++) {
                LineChartSeries bcs = chartSeries.ElementAt(i);

                string formulaVal = string.Format("Data!$B${0}:$" + lastCol + "${0}", (i + 2));
                DocumentFormat.OpenXml.Drawing.Charts.Values values = bcs.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Values());

                NumberReference numberReference = values.AppendChild(new NumberReference() {
                    Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formulaVal }
                });

                NumberingCache numberingCache = numberReference.AppendChild(new NumberingCache());
                numberingCache.Append(new PointCount() { Val = (uint)(dataTable.Columns.Count - 1) });

                for (int j = 1; j < dataTable.Columns.Count; j++) {
                    numberingCache.AppendChild(new NumericPoint() { Index = (uint)j }).Append(new NumericValue(dataTable.Rows[i][j].ToString()));
                }

            }

            lineChart.AppendChild(new DataLabels(
                                new ShowLegendKey() { Val = true },
                                new ShowValue() { Val = true },
                                new ShowCategoryName() { Val = false },
                                new ShowSeriesName() { Val = false },
                                new ShowPercent() { Val = false },
                                new ShowBubbleSize() { Val = false }
                            ));

            lineChart.Append(new AxisId() { Val = 48650112u });
            lineChart.Append(new AxisId() { Val = 48672768u });

            // Adding Category Axis
            plotArea.AppendChild(
                new CategoryAxis(
                    new AxisId() { Val = 48650112u },
                    new Scaling(new Orientation() { Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax) }),
                    new Delete() { Val = false },
                    new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
                    new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                    new CrossingAxis() { Val = 48672768u },
                    new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                    new AutoLabeled() { Val = true },
                    new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) }
                ));

            // Adding Value Axis
            plotArea.AppendChild(
                new ValueAxis(
                    new AxisId() { Val = 48672768u },
                    new Scaling(new Orientation() { Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax) }),
                    new Delete() { Val = false },
                    new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
                    new MajorGridlines(),
                    new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat() {
                        FormatCode = "General",
                        SourceLinked = true
                    },
                    new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                    new CrossingAxis() { Val = 48650112u },
                    new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                    new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }
                ));

            chart.Append(
                    new PlotVisibleOnly() { Val = true },
                    new DisplayBlanksAs() { Val = new EnumValue<DisplayBlanksAsValues>(DisplayBlanksAsValues.Gap) },
                    new ShowDataLabelsOverMaximum() { Val = false }
                );

            chartPart.ChartSpace.Save();

            // Positioning the chart on the spreadsheet
            TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild(new TwoCellAnchor());

            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(
                    new ColumnId("0"),
                    new ColumnOffset("0"),
                    new RowId((1).ToString()),
                    new RowOffset("0")
                ));

            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(
                    new ColumnId("20"),
                    new ColumnOffset("0"),
                    new RowId((35).ToString()),
                    new RowOffset("0")
                ));

            // Append GraphicFrame to TwoCellAnchor
            GraphicFrame graphicFrame = twoCellAnchor.AppendChild(new GraphicFrame());
            graphicFrame.Macro = string.Empty;

            graphicFrame.Append(new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties() {
                        Id = 2u,
                        Name = "Reget Statistics"
                    },
                    new NonVisualGraphicFrameDrawingProperties()
                ));

            graphicFrame.Append(new Transform(
                    new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                    new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 0L, Cy = 0L }
                ));

            graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Graphic(
                    new DocumentFormat.OpenXml.Drawing.GraphicData(
                            new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) }
                        ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                ));

            twoCellAnchor.Append(new ClientData());

            AddChartTitle(chart, chartTitle);
            #endregion

            xlsDoc.Close();

            return GetWorkbookMemoryStream();
        }

        private MemoryStream GenerateExcelChartPie(System.Data.DataTable dataTable, string chartTitle) {
            List<double> columnWidths = new List<double>();
            columnWidths.Add(50);
            for (int i = 1; i < dataTable.Columns.Count; i++) {
                columnWidths.Add(30);
            }

            SpreadsheetDocument xlsDoc = GenerateExcelWorkbookDoc(dataTable, columnWidths);

            string sheetName = "Chart";
            var wbPart = xlsDoc.WorkbookPart;
            var wsPart = AddWorkSheet(wbPart, sheetName, 0);

            #region Charts
            // Add a new drawing to the worksheet
            DrawingsPart drawingsPart = wsPart.AddNewPart<DrawingsPart>();
            wsPart.Worksheet.Append(new Drawing() { Id = wsPart.GetIdOfPart(drawingsPart) });
            wsPart.Worksheet.Save();

            drawingsPart.WorksheetDrawing = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
            // Add a new chart and set the chart language
            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new ChartSpace();
            chartPart.ChartSpace.AppendChild(new EditingLanguage() { Val = "en-US" });
            Chart chart = chartPart.ChartSpace.AppendChild(new Chart());


            // Create a new Clustered Column Chart
            PlotArea plotArea = chart.AppendChild(new PlotArea());
            Layout layout = plotArea.AppendChild(new Layout());

            PieChart pieChart = plotArea.AppendChild(new PieChart());

            
            //Legend legend = new Legend();
            //legend.

            //// Create chart series
            //string lastCol = CommonExcel.GetLetterFromColIndex(dataTable.Columns.Count);
            //for (int i = 1; i < dataTable.Columns.Count; i++) {
            //    PieChartSeries pieChartSeries = pieChart.AppendChild(new PieChartSeries(
            //        new Index() { Val = (uint)i },
            //        new Order() { Val = (uint)i },
            //        new SeriesText(new NumericValue() { Text = dataTable.Columns[i].ColumnName.ToString() })
            //    ));

            //    // Adding category axis to the chart
            //    CategoryAxisData categoryAxisData = pieChartSeries.AppendChild(new CategoryAxisData());

            //    // Category
            //    // Constructing the chart category
            //    string colLetter = GetColumnAddress(i + 1);
            //    string formulaCat = string.Format("Data!${0}$1:${0}$1", colLetter);

            //    StringReference stringReference = categoryAxisData.AppendChild(new StringReference() {
            //        Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formulaCat }
            //    });

            //    StringCache stringCache = stringReference.AppendChild(new StringCache());
            //    stringCache.Append(new PointCount() { Val = (uint)1 });

            //    stringCache.AppendChild(new NumericPoint() { Index = (uint)0 }).Append(new NumericValue(dataTable.Columns[i].ColumnName));

            //}

            // Create chart series
            string lastCol = CommonExcel.GetLetterFromColIndex(dataTable.Columns.Count);
            for (int i = 0; i < dataTable.Rows.Count; i++) {
                PieChartSeries lineChartSeries = pieChart.AppendChild(new PieChartSeries(
                    new Index() { Val = (uint)i },
                    new Order() { Val = (uint)i },
                    new SeriesText(new NumericValue() { Text = dataTable.Rows[i][0].ToString() })
                ));

                // Adding category axis to the chart
                CategoryAxisData categoryAxisData = lineChartSeries.AppendChild(new CategoryAxisData());

                // Category
                // Constructing the chart category
                string formulaCat = "Data!$B$1:$" + lastCol + "$1";

                StringReference stringReference = categoryAxisData.AppendChild(new StringReference() {
                    Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formulaCat }
                });

                StringCache stringCache = stringReference.AppendChild(new StringCache());
                stringCache.Append(new PointCount() { Val = (uint)(dataTable.Columns.Count - 1) });

                for (int j = 1; j < (dataTable.Columns.Count); j++) {
                    stringCache.AppendChild(new NumericPoint() { Index = (uint)j }).Append(new NumericValue(dataTable.Columns[j].ColumnName));
                }
            }



            var chartSeries = pieChart.Elements<PieChartSeries>();

            //for (int i = 1; i < dataTable.Columns.Count; i++) {
            //    PieChartSeries bcs = chartSeries.ElementAt(i-1);

            //    string colLetter = GetColumnAddress(i + 1);
            //    string formulaVal = string.Format("Data!${0}$2:${0}$2", colLetter);

            //    DocumentFormat.OpenXml.Drawing.Charts.Values values = bcs.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Values());

            //    NumberReference numberReference = values.AppendChild(new NumberReference() {
            //        Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formulaVal }
            //    });

            //    NumberingCache numberingCache = numberReference.AppendChild(new NumberingCache());
            //    numberingCache.Append(new PointCount() { Val = (uint)1 });

            //    numberingCache.AppendChild(new NumericPoint() { Index = (uint)0 }).Append(new NumericValue(dataTable.Rows[0][i].ToString()));

            //}

            for (int i = 0; i < dataTable.Rows.Count; i++) {
                PieChartSeries bcs = chartSeries.ElementAt(i);

                string formulaVal = string.Format("Data!$B${0}:$" + lastCol + "${0}", (i + 2));
                DocumentFormat.OpenXml.Drawing.Charts.Values values = bcs.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Values());

                NumberReference numberReference = values.AppendChild(new NumberReference() {
                    Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formulaVal }
                });

                NumberingCache numberingCache = numberReference.AppendChild(new NumberingCache());
                numberingCache.Append(new PointCount() { Val = (uint)(dataTable.Columns.Count - 1) });

                for (int j = 1; j < dataTable.Columns.Count; j++) {
                    numberingCache.AppendChild(new NumericPoint() { Index = (uint)j }).Append(new NumericValue(dataTable.Rows[i][j].ToString()));
                }

            }

            pieChart.AppendChild(new DataLabels(
                                new ShowLegendKey() { Val = true },
                                new ShowValue() { Val = true },
                                new ShowCategoryName() { Val = false },
                                new ShowSeriesName() { Val = false },
                                new ShowPercent() { Val = false },
                                new ShowBubbleSize() { Val = false }
                            ));

            //pieChart.Append(new AxisId() { Val = 48650112u });
            //pieChart.Append(new AxisId() { Val = 48672768u });

            //// Adding Category Axis
            //plotArea.AppendChild(
            //    new CategoryAxis(
            //        new AxisId() { Val = 48650112u },
            //        new Scaling(new Orientation() { Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax) }),
            //        new Delete() { Val = false },
            //        new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
            //        new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
            //        new CrossingAxis() { Val = 48672768u },
            //        new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
            //        new AutoLabeled() { Val = true },
            //        new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) }
            //    ));

            //// Adding Value Axis
            //plotArea.AppendChild(
            //    new ValueAxis(
            //        new AxisId() { Val = 48672768u },
            //        new Scaling(new Orientation() { Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax) }),
            //        new Delete() { Val = false },
            //        new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
            //        new MajorGridlines(),
            //        new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat() {
            //            FormatCode = "General",
            //            SourceLinked = true
            //        },
            //        new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
            //        new CrossingAxis() { Val = 48650112u },
            //        new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
            //        new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }
            //    ));

            chart.Append(
                    new PlotVisibleOnly() { Val = true },
                    new DisplayBlanksAs() { Val = new EnumValue<DisplayBlanksAsValues>(DisplayBlanksAsValues.Gap) },
                    new ShowDataLabelsOverMaximum() { Val = false }
                );

            chartPart.ChartSpace.Save();

            // Positioning the chart on the spreadsheet
            TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild(new TwoCellAnchor());

            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(
                    new ColumnId("0"),
                    new ColumnOffset("0"),
                    new RowId((1).ToString()),
                    new RowOffset("0")
                ));

            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(
                    new ColumnId("20"),
                    new ColumnOffset("0"),
                    new RowId((35).ToString()),
                    new RowOffset("0")
                ));

            // Append GraphicFrame to TwoCellAnchor
            GraphicFrame graphicFrame = twoCellAnchor.AppendChild(new GraphicFrame());
            graphicFrame.Macro = string.Empty;

            graphicFrame.Append(new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties() {
                        Id = 2u,
                        Name = "Reget Statistics"
                    },
                    new NonVisualGraphicFrameDrawingProperties()
                ));

            graphicFrame.Append(new Transform(
                    new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                    new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 0L, Cy = 0L }
                ));

            graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Graphic(
                    new DocumentFormat.OpenXml.Drawing.GraphicData(
                            new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) }
                        ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                ));

            twoCellAnchor.Append(new ClientData());

            AddChartTitle(chart, chartTitle);
            #endregion

            xlsDoc.Close();

            return GetWorkbookMemoryStream();
        }

        private void AddChartTitle(DocumentFormat.OpenXml.Drawing.Charts.Chart chart, string title) {
            var ctitle = chart.AppendChild(new Title());
            var chartText = ctitle.AppendChild(new ChartText());
            var richText = chartText.AppendChild(new RichText());

            var bodyPr = richText.AppendChild(new DocumentFormat.OpenXml.Drawing.BodyProperties());
            var lstStyle = richText.AppendChild(new DocumentFormat.OpenXml.Drawing.ListStyle());
            var paragraph = richText.AppendChild(new DocumentFormat.OpenXml.Drawing.Paragraph());

            var apPr = paragraph.AppendChild(new DocumentFormat.OpenXml.Drawing.ParagraphProperties());
            apPr.AppendChild(new DocumentFormat.OpenXml.Drawing.DefaultRunProperties());

            var run = paragraph.AppendChild(new DocumentFormat.OpenXml.Drawing.Run());
            run.AppendChild(new DocumentFormat.OpenXml.Drawing.RunProperties() { Language = "en-CA" });
            run.AppendChild(new DocumentFormat.OpenXml.Drawing.Text() { Text = title });
        }

        public static Column CreateColumnData(UInt32 StartColumnIndex, UInt32 EndColumnIndex, double ColumnWidth) {
            Column column;
            column = new Column();
            column.Min = StartColumnIndex;
            column.Max = EndColumnIndex;
            column.Width = ColumnWidth;
            column.CustomWidth = true;
            return column;
        }

        public static void SetAutoFilter(Worksheet ws, string strFilterRef) {
            AutoFilter autoFilter = new AutoFilter() { Reference = strFilterRef };
            ws.Append(autoFilter);
        }


        public Cell InsertCellInWorksheet(Worksheet ws, string addressName) {
            SheetData sheetData = ws.GetFirstChild<SheetData>();
            Cell cell = null;

            UInt32 rowNumber = GetRowIndex(addressName);
            Row row = GetRow(sheetData, rowNumber);

            // If the cell you need already exists, return it.
            // If there is not a cell with the specified column name, insert one.  
            Cell refCell = row.Elements<Cell>().
                Where(c => c.CellReference.Value == addressName).FirstOrDefault();
            if (refCell != null) {
                cell = refCell;
            } else {
                cell = CreateCell(row, addressName);
            }
            return cell;
        }

        /// <summary>
        /// Save time if a new document is generated
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="addressName"></param>
        /// <returns></returns>
        public Cell InsertNewCellInWorksheet(Worksheet ws, Row row, string addressName) {
            Cell cell = null;
                        
            cell = CreateCell(row, addressName);
            
            return cell;
        }

        //// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        //// If the cell already exists, returns it. 
        //private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart) {
        //    Worksheet worksheet = worksheetPart.Worksheet;
        //    SheetData sheetData = worksheet.GetFirstChild<SheetData>();
        //    string cellReference = columnName + rowIndex;

        //    // If the worksheet does not contain a row with the specified row index, insert one.
        //    Row row;
        //    if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0) {
        //        row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        //    } else {
        //        row = new Row() { RowIndex = rowIndex };
        //        sheetData.Append(row);
        //    }

        //    // If there is not a cell with the specified column name, insert one.  
        //    if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0) {
        //        return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
        //    } else {
        //        // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
        //        Cell refCell = null;
        //        foreach (Cell cell in row.Elements<Cell>()) {
        //            if (string.Compare(cell.CellReference.Value, cellReference, true) > 0) {
        //                refCell = cell;
        //                break;
        //            }
        //        }

        //        Cell newCell = new Cell() { CellReference = cellReference };
        //        row.InsertBefore(newCell, refCell);

        //        worksheet.Save();
        //        return newCell;
        //    }
        //}

        // Add a cell with the specified address to a row.
        private Cell CreateCell(Row row, String address) {
            Cell cellResult;
            Cell refCell = null;

            //************* code below causes problems if there are more then 25 columns , AA1, AB2
            //// Cells must be in sequential order according to CellReference. 
            //// Determine where to insert the new cell.
            //foreach (Cell cell in row.Elements<Cell>()) {
            //    if (string.Compare(cell.CellReference.Value, address, true) > 0) {
            //        refCell = cell;
            //        break;
            //    }
            //}
            //*************************************************************************************

            cellResult = new Cell();
            cellResult.CellReference = address;

            row.InsertBefore(cellResult, refCell);
            return cellResult;
        }

        private Row GetRow(SheetData wsData, UInt32 rowIndex) {
            var row = wsData.Elements<Row>().
            Where(r => r.RowIndex.Value == rowIndex).FirstOrDefault();
            if (row == null) {
                row = new Row();
                row.RowIndex = rowIndex;
                wsData.Append(row);
            }
            return row;
        }

        // Given an Excel address such as E5 or AB128, GetRowIndex
        // parses the address and returns the row index.
        private static UInt32 GetRowIndex(string address) {
            string rowPart;
            UInt32 l;
            UInt32 result = 0;

            for (int i = 0; i < address.Length; i++) {
                if (UInt32.TryParse(address.Substring(i, 1), out l)) {
                    rowPart = address.Substring(i, address.Length - i);
                    if (UInt32.TryParse(rowPart, out l)) {
                        result = l;
                        break;
                    }
                }
            }
            return result;
        }

        //private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart) {
        //    // If the part does not contain a SharedStringTable, create one.
        //    if (shareStringPart.SharedStringTable == null) {
        //        shareStringPart.SharedStringTable = new SharedStringTable();
        //    }

        //    int i = 0;

        //    // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        //    foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>()) {
        //        if (item.InnerText == text) {
        //            return i;
        //        }

        //        i++;
        //    }

        //    // The text does not exist in the part. Create the SharedStringItem and return its index.
        //    shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
        //    shareStringPart.SharedStringTable.Save();

        //    return i;
        //}


        public static string GetColumnAddress(int columnNumber) {
            int dividend = columnNumber;
            string columnName = "";
            int modulo;

            while (dividend > 0) {
                modulo = (dividend - 1) % 26;
                columnName = ((char)(65 + modulo)).ToString() + columnName;
                dividend = ((dividend - modulo) / 26);
                
            }
            return columnName;

        }

        public void SetCellValue(Cell cell, object oValue) {
            if (oValue == null) {
                cell.CellValue = new CellValue("");
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
                return;
            }

            string strValue = oValue.ToString();

            //strValue = strValue.Replace('\u0006', "");
            //if (strValue.Contains('\u0006')) {
            //    int h = 5;
            //}

            Type t = oValue.GetType();
            if (t == typeof(string)) {
                cell.CellValue = new CellValue(strValue);
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
            } else if(t== typeof(Int16) || t == typeof(Int32) || t == typeof(Decimal) || t == typeof(Double)) {
                cell.CellValue = new CellValue(strValue);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            } else {
                cell.CellValue = new CellValue(strValue);
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
            }

            
        }

        public WorksheetPart AddSheet(WorkbookPart wbPart, string sheetName) {
            sheetName = sheetName.Replace(@"\", "_").Replace("/", "_").Replace("*", "_").Replace("[", "_").Replace("]", "_").Replace(":", "_").Replace("?", "_");
            if (sheetName.Length > 25) {
                sheetName = sheetName.Substring(0, 25) + "...";
            }
            var wsPart = AddWorkSheet(wbPart, sheetName);

            return wsPart;
        }

        

        //public static SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id) {
        //    return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        //}

        //public static object GetCellValue(Cell cell, WorkbookPart wbPart) {
        //    if (cell.DataType != null) {

        //        if (cell.DataType == CellValues.SharedString) {
        //            int id = -1;

        //            if (Int32.TryParse(cell.InnerText, out id)) {
        //                SharedStringItem item = GetSharedStringItemById(wbPart, id);

        //                if (item.Text != null) {
        //                    //code to take the string value  
        //                    return item.Text.Text;
        //                } else if (item.InnerText != null) {
        //                    return item.InnerText;
        //                } else if (item.InnerXml != null) {
        //                    return item.InnerXml;
        //                }
        //            }
        //        } else if (cell.DataType == CellValues.Boolean) {
        //            switch (cell.InnerText) {
        //                case "0":
        //                    return "false";
        //                default:
        //                    return "true"; 
        //            }
        //        }
        //    } else {
        //       // CellFormat cellFormat = (CellFormat)wbPart.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt(Convert.ToInt32(cell.StyleIndex.Value));
        //        //    string format = wbPart.WorkbookStylesPart.Stylesheet.NumberingFormats.Elements<NumberingFormat>()
        //        //.Where(i => i.NumberFormatId.ToString() == cellFormat.NumberFormatId.ToString())
        //        //.First().FormatCode;
        //        //return cell.CellValue;

        //        return GetCellValueWithoutConsideringDataType(wbPart, cell);
        //    }

        //    return null;
        //}

        //public static object GetCellText(Cell cell, WorkbookPart wbPart) {
        //    if (cell.DataType != null) {

        //        if (cell.DataType == CellValues.SharedString) {
        //            int id = -1;

        //            if (Int32.TryParse(cell.InnerText, out id)) {
        //                SharedStringItem item = GetSharedStringItemById(wbPart, id);

        //                if (item.Text != null) {
        //                    //code to take the string value  
        //                    return item.Text.Text;
        //                } else if (item.InnerText != null) {
        //                    return item.InnerText;
        //                } else if (item.InnerXml != null) {
        //                    return item.InnerXml;
        //                }
        //            }
        //        } else {
        //            return cell.CellValue.InnerText;
        //        }
        //    } else {
        //        return cell.CellValue.InnerText;
        //    }

        //    return null;
        //}

        //private static object GetCellValueWithoutConsideringDataType(WorkbookPart workbookPart, Cell excelCell) {
        //    CellFormat cellFormat = GetCellFormat(workbookPart, excelCell);
        //    if (cellFormat != null) {
        //        return GetFormatedValue(workbookPart, excelCell, cellFormat);
        //    } else {
        //        return excelCell.InnerText;
        //    }
        //}

        private static CellFormat GetCellFormat(WorkbookPart workbookPart, Cell cell) {
            //WorkbookPart workbookPart = GetWorkbookPartFromCell(cell);
            int styleIndex = (int)cell.StyleIndex.Value;
            CellFormat cellFormat = (CellFormat)workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt(styleIndex);
            return cellFormat;
        }

        public static string Unlock(string filePath) {
            FileInfo fi = new FileInfo(filePath);

            
            string zipFileName = GetOrigFileName(fi.Directory.FullName, System.IO.Path.GetFileNameWithoutExtension(fi.Name),  ".zip");
            File.Copy(fi.FullName, zipFileName);
            
            //create zip folder
            FileInfo fiZip = new FileInfo(zipFileName);
            string zipFolder = System.IO.Path.Combine(fi.Directory.FullName, System.IO.Path.ChangeExtension(fiZip.Name, ""));
            zipFolder = zipFolder.Substring(0, zipFolder.Length - 1);
            zipFolder = GetOrigDirectoryName(zipFolder);

            try {
                ZipFile.ExtractToDirectory(zipFileName, zipFolder);
                File.Delete(zipFileName);
                string xlFolder = System.IO.Path.Combine(zipFolder, "xl");
                string xlsSheetFolder = System.IO.Path.Combine(xlFolder, "worksheets");

                foreach (var ws in Directory.GetFiles(xlsSheetFolder)) {
                    RemoveSheetProtection(ws);
                }

                string xlVba = Path.Combine(xlFolder, "vbaProject.bin");
                if (File.Exists(xlVba)) {
                    RemoveVbaPasswordProtection(xlVba);
                }


                string unlockFileName = GetOrigFileName(fi.Directory.FullName, System.IO.Path.GetFileNameWithoutExtension(fi.Name) + "_unlock", System.IO.Path.GetExtension(fi.Name));
                ZipFile.CreateFromDirectory(zipFolder, unlockFileName);
                return unlockFileName;
            } catch (Exception ex) {
                throw ex;
            } finally {
                DeleteDirectory(zipFolder);
            }

        }

        private static void RemoveSheetProtection(string lockFileName) {
            FileInfo fi = new FileInfo(lockFileName);
            string unlockFileName = lockFileName + ".fix";

            using (StreamWriter sw = new StreamWriter(unlockFileName)) {
                IEnumerable<string> lines = File.ReadLines(lockFileName);

                foreach (var line in lines) {
                    bool isFixed = false;
                    int iPosStart = line.IndexOf("<sheetProtection");
                    if (iPosStart >= 0) {
                        int iPosEnd = line.IndexOf(">", iPosStart + 1);
                        if (iPosEnd > -1) {
                            string strFixedLine = line.Substring(0, iPosStart);
                            strFixedLine += line.Substring(iPosEnd + 1);
                            sw.WriteLine(strFixedLine);
                            isFixed = true;
                        }
                    }

                    if (!isFixed) {
                        sw.WriteLine(line);
                    }

                    
                }
            }

            File.Delete(lockFileName);
            File.Move(unlockFileName, lockFileName);
        }

        private static void RemoveVbaPasswordProtection(string lockFileName) {


            FileInfo fi = new FileInfo(lockFileName);
            string unlockFileName = lockFileName + ".fix";

            //search for DPB="
            long posDbpStart = FindBytesPosition(lockFileName, new int[] { 68, 80, 66, 61, 34 });
            //long posDbpStop = FindBytesPosition(lockFileName, new int[] { 34 }, posDbpStart + 1);

            long lPpos = posDbpStart - 2;

            using (FileStream fsw = new FileStream(unlockFileName, FileMode.CreateNew)) {
                using (FileStream fsr = new FileStream(lockFileName, FileMode.Open)) {
                    int iChar;
                    while ((iChar = fsr.ReadByte()) != -1) {

                        if (fsr.Position == lPpos) {
                            fsw.WriteByte(Convert.ToByte(68));
                        } else {
                            fsw.WriteByte(Convert.ToByte(iChar));

                        }
                    }
                }
            }

            //using (FileStream fs = new FileStream(lockFileName, FileMode.Open)) {


            //    //int hexIn;
            //    //String hex;

            //    //for (int i = 0; (hexIn = fs.ReadByte()) != -1; i++) {
            //    //    hex = string.Format("{0:X2}", hexIn);
            //    //}
            //}

            File.Delete(lockFileName);
            File.Move(unlockFileName, lockFileName);
        }

        private static void RemoveVbaPasswordProtectionReplace(string lockFileName) {


            FileInfo fi = new FileInfo(lockFileName);
            string unlockFileName = lockFileName + ".fix";

            //search for DPB="
            long posDbpStart = FindBytesPosition(lockFileName, new int[] { 68, 80, 66, 61, 34 });
            long posDbpStop = FindBytesPosition(lockFileName, new int[] { 34 }, posDbpStart + 1);

            string refVbProject = @"C:\Temp\XlsTest\vbawopassword\xl\vbaProject.bin";
            long refPosDbpStart = FindBytesPosition(refVbProject, new int[] { 68, 80, 66, 61, 34 });
            long refPosDbpStop = FindBytesPosition(refVbProject, new int[] { 34 }, refPosDbpStart + 1);


            using (FileStream fsw = new FileStream(unlockFileName, FileMode.CreateNew)) {
                using (FileStream fsr = new FileStream(lockFileName, FileMode.Open)) {

                    int iChar;
                    while ((iChar = fsr.ReadByte()) != -1) {

                        if (fsr.Position <= posDbpStart) {
                            fsw.WriteByte(Convert.ToByte(iChar));
                        } else {
                            break;

                        }
                    }
                }

            }

            using (FileStream fsw = new FileStream(unlockFileName, FileMode.Append)) {
                using (FileStream fsrRef = new FileStream(refVbProject, FileMode.Open)) {

                    fsrRef.Seek(refPosDbpStart, SeekOrigin.Begin);
                    int iChar;
                    while ((iChar = fsrRef.ReadByte()) != -1) {

                        if (fsrRef.Position <= refPosDbpStop) {
                            fsw.WriteByte(Convert.ToByte(iChar));
                        } else {
                            break;

                        }
                    }
                }
            }

            using (FileStream fsw = new FileStream(unlockFileName, FileMode.Append)) {
                using (FileStream fsr = new FileStream(lockFileName, FileMode.Open)) {
                    fsr.Seek(posDbpStop, SeekOrigin.Begin);
                    int iChar;
                    while ((iChar = fsr.ReadByte()) != -1) {


                        fsw.WriteByte(Convert.ToByte(iChar));

                    }
                }

            }

            //using (FileStream fs = new FileStream(lockFileName, FileMode.Open)) {


            //    //int hexIn;
            //    //String hex;

            //    //for (int i = 0; (hexIn = fs.ReadByte()) != -1; i++) {
            //    //    hex = string.Format("{0:X2}", hexIn);
            //    //}
            //}

            File.Delete(lockFileName);
            File.Move(unlockFileName, lockFileName);
        }

        private static long FindBytesPosition(string lockFileName, int[] searchBytes) {
            return FindBytesPosition(lockFileName, searchBytes, 0);
        }

        private static long FindBytesPosition(string lockFileName, int[] searchBytes, long startPos) {
            using (FileStream fs = new FileStream(lockFileName, FileMode.Open)) {
                int iChar;
                int iSearchBytesIndex = 0;

                if (startPos > 0) {
                    var tt = fs.Length;
                    fs.Seek(startPos, SeekOrigin.Begin);
                }

                while ((iChar = fs.ReadByte()) != -1) {
                    if (iChar == searchBytes[iSearchBytesIndex]) {
                        iSearchBytesIndex++;
                        if (searchBytes.Length == 1 || iSearchBytesIndex == searchBytes.Length - 1) {
                            return fs.Position;
                        }
                    } else {
                        iSearchBytesIndex = 0;
                    }
                }
            }

            return -1;
        }

        private static string GetOrigFileName(string folder, string pureFileName, string extention) {
            string fileName = System.IO.Path.Combine(folder, pureFileName + extention);
            int iIndex = 1;
            while (File.Exists(fileName)) {
                fileName = System.IO.Path.Combine(folder, pureFileName + "_" + iIndex + extention);
                iIndex++;
            }

            return fileName;
        }

        private static string GetOrigDirectoryName(string folder) {
            int iIndex = 1;
            string pureFolder = folder;
            string tmpFolder = folder;
            while (Directory.Exists(tmpFolder)) {
                tmpFolder = pureFolder + "_" + iIndex;
                iIndex++;
            }

            return tmpFolder;
        }

        private static void DeleteDirectory(string strPath) {
            foreach (var fileName in Directory.GetFiles(strPath)) {
                File.Delete(fileName);
            }

            foreach (var childDir in Directory.GetDirectories(strPath)) {
                DeleteDirectory(childDir);
            }

            Directory.Delete(strPath);
        }

        // Given a document name, a worksheet name, and the names of two adjacent cells, merges the two cells.
        // When two cells are merged, only the content from one cell is preserved:
        // the upper-left cell for left-to-right languages or the upper-right cell for right-to-left languages.
        public static void MergeTwoCells(Worksheet worksheet, string cell1Name, string cell2Name) {
            // Open the document for editing.
            
                
                if (worksheet == null || string.IsNullOrEmpty(cell1Name) || string.IsNullOrEmpty(cell2Name)) {
                    return;
                }

                // Verify if the specified cells exist, and if they do not exist, create them.
                CreateSpreadsheetCellIfNotExist(worksheet, cell1Name);
                CreateSpreadsheetCellIfNotExist(worksheet, cell2Name);

                MergeCells mergeCells;
                if (worksheet.Elements<MergeCells>().Count() > 0) {
                    mergeCells = worksheet.Elements<MergeCells>().First();
                } else {
                    mergeCells = new MergeCells();

                    // Insert a MergeCells object into the specified position.
                    if (worksheet.Elements<CustomSheetView>().Count() > 0) {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
                    } else if (worksheet.Elements<DataConsolidate>().Count() > 0) {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
                    } else if (worksheet.Elements<SortState>().Count() > 0) {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
                    } else if (worksheet.Elements<AutoFilter>().Count() > 0) {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
                    } else if (worksheet.Elements<Scenarios>().Count() > 0) {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
                    } else if (worksheet.Elements<ProtectedRanges>().Count() > 0) {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
                    } else if (worksheet.Elements<SheetProtection>().Count() > 0) {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
                    } else if (worksheet.Elements<SheetCalculationProperties>().Count() > 0) {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
                    } else {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
                    }
                }

                // Create the merged cell and append it to the MergeCells collection.
                MergeCell mergeCell = new MergeCell() { Reference = new StringValue(cell1Name + ":" + cell2Name) };
                mergeCells.Append(mergeCell);

                worksheet.Save();
            
        }
        // Given a Worksheet and a cell name, verifies that the specified cell exists.
        // If it does not exist, creates a new cell. 
        private static void CreateSpreadsheetCellIfNotExist(Worksheet worksheet, string cellName) {
            string columnName = GetColumnName(cellName);
            uint rowIndex = GetRowIndex(cellName);

            IEnumerable<Row> rows = worksheet.Descendants<Row>().Where(r => r.RowIndex.Value == rowIndex);

            // If the Worksheet does not contain the specified row, create the specified row.
            // Create the specified cell in that row, and insert the row into the Worksheet.
            if (rows.Count() == 0) {
                Row row = new Row() { RowIndex = new UInt32Value(rowIndex) };
                Cell cell = new Cell() { CellReference = new StringValue(cellName) };
                row.Append(cell);
                worksheet.Descendants<SheetData>().First().Append(row);
                worksheet.Save();
            } else {
                Row row = rows.First();

                IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellName);

                // If the row does not contain the specified cell, create the specified cell.
                if (cells.Count() == 0) {
                    Cell cell = new Cell() { CellReference = new StringValue(cellName) };
                    row.Append(cell);
                    worksheet.Save();
                }
            }
        }

        // Given a cell name, parses the specified cell to get the column name.
        private static string GetColumnName(string cellName) {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }
        // Given a cell name, parses the specified cell to get the row index.
        //private static uint GetRowIndex(string cellName) {
        //    // Create a regular expression to match the row index portion the cell name.
        //    Regex regex = new Regex(@"\d+");
        //    Match match = regex.Match(cellName);

        //    return uint.Parse(match.Value);
        //}

        //private static WorkbookPart GetWorkbookPartFromCell(Cell cell) {
        //    Worksheet workSheet = cell.Ancestors<Worksheet>().FirstOrDefault();
        //    SpreadsheetDocument doc = workSheet.WorksheetPart.OpenXmlPackage as SpreadsheetDocument;
        //    return doc.WorkbookPart;
        //}

        //private static object GetFormatedValue(WorkbookPart workbookPart, Cell cell, CellFormat cellformat) {
        //    //string value;

        //    if (cellformat.NumberFormatId.Value != 0) {
        //        int iNumFormatId = (int)cellformat.NumberFormatId.Value;
        //        if (numberFormatDictionary.ContainsKey(iNumFormatId)) {
        //            string strFormat = numberFormatDictionary[iNumFormatId];
        //            switch(iNumFormatId) {
        //                case 1:
        //                case 2:
        //                case 3:
        //                case 4:
        //                    return double.Parse(cell.InnerText);
        //                default:
        //                    return cell.InnerText;
        //            }
        //        } else if (workbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats != null) {
        //            var numberFormat = workbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats.Elements<NumberingFormat>()
        //                .Where(i => i.NumberFormatId.Value == cellformat.NumberFormatId.Value)
        //                .FirstOrDefault();

        //            string strFormat = numberFormat.FormatCode;
        //            double number = double.Parse(cell.InnerText);
        //            return number.ToString(strFormat);
        //        }


        //    } else {
        //        return cell.InnerText;
        //    }
        //    return cell.InnerText;
        //}

        //private static Dictionary<int, string> numberFormatDictionary = new Dictionary<int, string>()
        //    {
        //        {0, "General"},
        //        {1, "0"},
        //        {2, "0.00"},
        //        {3, "#,##0"},
        //        {4, "#,##0.00"},
        //        {9, "0%"},
        //        {10, "0.00%"},
        //        {11, "0.00E+00"},
        //        {12, "# ?/?"},
        //        {13, "# ??/??"},
        //        {14, "mm-dd-yy"},
        //        {15, "d-mmm-yy"},
        //        {16, "d-mmm"},
        //        {17, "mmm-yy"},
        //        {18, "h:mm AM/PM"},
        //        {19, "h:mm:ss AM/PM"},
        //        {20, "h:mm"},
        //        {21, "h:mm:ss"},
        //        {22, "m/d/yy h:mm"},
        //        {37, "#,##0 ;(#,##0)"},
        //        {38, "#,##0 ;[Red](#,##0)"},
        //        {39, "#,##0.00;(#,##0.00)"},
        //        {40, "#,##0.00;[Red](#,##0.00)"},
        //        {45, "mm:ss"},
        //        {46, "[h]:mm:ss"},
        //        {47, "mmss.0"},
        //        {48, "##0.0E+0"},
        //        {49, "@"}
        //    };
        #endregion
    }
}
