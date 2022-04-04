using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kamsyk.ExcelOpenXml.ExcelActiveX {
    public class Range : BaseNotInstantiable {
        #region Properties
        //private List<DocumentFormat.OpenXml.Spreadsheet.Cell> _m_XmlCells = null;
        //private List<DocumentFormat.OpenXml.Spreadsheet.Cell> m_XmlCells {
        //    get {
        //        if (_m_XmlCells == null) {
        //            _m_XmlCells = GetXmlCells();
        //            //_m_XmlCells = new List<DocumentFormat.OpenXml.Spreadsheet.Cell>();
        //            //for (int iRow = m_RowIndexStart; iRow <= m_RowIndexEnd; iRow++) {
        //            //    var xmlRow = m_Worksheet.GetRow(iRow);
        //            //    for (int j = m_ColIndexStart; j <= m_ColIndexEnd; j++) {
        //            //        if (xmlRow != null) {
        //            //            var xmlCell = m_Worksheet.GetCell(xmlRow, j);
        //            //            if (xmlCell == null) {
        //            //                xmlCell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
        //            //            }
        //            //            _m_XmlCells.Add(xmlCell);

        //            //        } else {
        //            //            DocumentFormat.OpenXml.Spreadsheet.Cell xmlCell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
        //            //            _m_XmlCells.Add(xmlCell);
        //            //        }
        //            //    }
        //            //}
        //        }

        //        return _m_XmlCells;
        //    }
        //}

        private Worksheet m_Worksheet = null;

        public object Value {
            get { return GetRangeValue(); }
        }

        public object Text {
            get { return GetRangeValue(); }
        }

        private int m_RowIndexStart = -1;
        private int m_ColIndexStart = -1;
        private int m_RowIndexEnd = -1;
        private int m_ColIndexEnd = -1;

        
        //private List<Cell> m_Cells = null;
        //public List<Cell> Cells {
        //    get {
        //        if (m_Cells == null) {
        //            m_Cells = GetCells();
        //        }

        //        return m_Cells;
        //    }
        //}
        #endregion

        #region Disctionary
        private static Dictionary<int, string> numberFormatDictionary = new Dictionary<int, string>()
            {
                {0, "General"},
                {1, "0"},
                {2, "0.00"},
                {3, "#,##0"},
                {4, "#,##0.00"},
                {9, "0%"},
                {10, "0.00%"},
                {11, "0.00E+00"},
                {12, "# ?/?"},
                {13, "# ??/??"},
                {14, "mm-dd-yy"},
                {15, "d-mmm-yy"},
                {16, "d-mmm"},
                {17, "mmm-yy"},
                {18, "h:mm AM/PM"},
                {19, "h:mm:ss AM/PM"},
                {20, "h:mm"},
                {21, "h:mm:ss"},
                {22, "m/d/yy h:mm"},
                {37, "#,##0 ;(#,##0)"},
                {38, "#,##0 ;[Red](#,##0)"},
                {39, "#,##0.00;(#,##0.00)"},
                {40, "#,##0.00;[Red](#,##0.00)"},
                {45, "mm:ss"},
                {46, "[h]:mm:ss"},
                {47, "mmss.0"},
                {48, "##0.0E+0"},
                {49, "@"}
            };
        #endregion

        #region Constructor
        internal Range(Worksheet worksheet, int iRowStart, int iColStart, int iRowEnd, int iColEnd) {
            m_Worksheet = worksheet;
            m_RowIndexStart = iRowStart;
            m_ColIndexStart = iColStart;
            m_RowIndexEnd = iRowEnd;
            m_ColIndexEnd = iColEnd;
        }
        #endregion

        #region Static Methods
        private static object GetCellFormatValue(string strValue, int iNumFormatId) {
            switch (iNumFormatId) {
                case 1:
                case 2:
                case 3:
                case 4:
                    return double.Parse(strValue);
                case 14:
                case 15:
                case 16:
                case 17:
                    int iDate = int.Parse(strValue);
                    DateTime dDate = Excel.ZERO_DATE.AddDays(iDate);
                    return dDate;
                default:
                    return strValue;
            }
        }
        #endregion

        #region Methods
        private object GetRangeValue() {
            //so called SAX approach
            List<object> retValues = new List<object>();

            int iRowCount = m_RowIndexEnd - m_RowIndexStart;
            if (iRowCount < 0) {
                int iTmp = m_RowIndexEnd;
                m_RowIndexEnd = m_RowIndexStart;
                m_RowIndexStart = iTmp;
            }


            int iColCount = m_ColIndexEnd - m_ColIndexStart;
            if (iColCount < 0) {
                int iTmp = m_ColIndexEnd;
                m_ColIndexEnd = m_ColIndexStart;
                m_ColIndexStart = iTmp;
            }

            Hashtable htSharedTexts = new Hashtable();

            //DocumentFormat.OpenXml.Packaging.WorksheetPart wsPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)CurrWorkbookPart.GetPartById(m_Worksheet.XmlSheet.Id);
            //DocumentFormat.OpenXml.OpenXmlReader readerXml = DocumentFormat.OpenXml.OpenXmlReader.Create(wsPart);
            DocumentFormat.OpenXml.OpenXmlReader readerXml = DocumentFormat.OpenXml.OpenXmlReader.Create(m_Worksheet.WorksheetPart);

            try {
                CellFormat[] cellFormats = null;
                if (CurrWorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats != null && 
                    CurrWorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>() != null) {
                    cellFormats = CurrWorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ToArray();
                }

                int iLastRow = m_RowIndexStart - 1;
                while (readerXml.Read()) {
                    if (readerXml.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.Row)) {
                        bool isRowInRange = false;
                        string strRowIndex = null;
                        int iCurrRowIndex = -1;
                        if (readerXml.HasAttributes) {
                            strRowIndex = readerXml.Attributes.First(a => a.LocalName == "r").Value;
                            iCurrRowIndex = Convert.ToInt32(strRowIndex);
                            if (m_RowIndexEnd < iCurrRowIndex) {
                                break;
                            }
                            isRowInRange = (m_RowIndexStart <= iCurrRowIndex);
                        }

                        if (isRowInRange) {
                            for (int i = iLastRow + 1; i < iCurrRowIndex; i++) {
                                for (int j = m_ColIndexStart; j <= m_ColIndexEnd; j++) {
                                    retValues.Add(null);
                                }
                            }
                            iLastRow = iCurrRowIndex;

                            GetRangeRowValues(
                                readerXml,
                                strRowIndex,
                                ref htSharedTexts,
                                cellFormats,
                                ref retValues);

                            if (iCurrRowIndex == m_RowIndexEnd) {
                                break;
                            }

                            //int iLastCol = m_ColIndexStart - 1;
                            //readerXml.ReadFirstChild();
                            //do {

                            //    if (readerXml.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.Cell)) {
                            //        bool isColInRange = false;
                            //        int iCurrColIndex = -1;
                            //        string strCellFormat = null;
                            //        if (readerXml.HasAttributes) {
                            //            string cellRef = readerXml.Attributes.First(a => a.LocalName == "r").Value;
                            //            var xmlCellFormat = readerXml.Attributes.SingleOrDefault(attr => attr.LocalName == "s");
                            //            if (xmlCellFormat != null) {
                            //                strCellFormat = xmlCellFormat.Value;
                            //            }
                            //            string strCol = cellRef.Replace(strRowIndex, "");

                            //            iCurrColIndex = Worksheet.GetColIndexFromLetter(strCol);
                            //            if (m_ColIndexEnd < iCurrColIndex) {
                            //                break;
                            //            }
                            //            isColInRange = (m_ColIndexStart <= iCurrColIndex);
                            //        }

                            //        if (isColInRange) {
                            //            for (int i = iLastCol + 1; i < iCurrColIndex; i++) {
                            //                retValues.Add(null);
                            //            }
                            //            iLastCol = iCurrColIndex;

                            //            DocumentFormat.OpenXml.Spreadsheet.Cell c = (DocumentFormat.OpenXml.Spreadsheet.Cell)readerXml.LoadCurrentElement();

                            //            object cellValue = null;

                            //            if (c.CellValue != null) {
                            //                if (c.DataType != null && c.DataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) {
                            //                    if (htSharedTexts.ContainsKey(c.CellValue.InnerText)) {
                            //                        cellValue = htSharedTexts[c.CellValue.InnerText].ToString();
                            //                    } else {
                            //                        DocumentFormat.OpenXml.Spreadsheet.SharedStringItem ssi = CurrWorkbookPart.SharedStringTablePart.SharedStringTable.Elements<DocumentFormat.OpenXml.Spreadsheet.SharedStringItem>().ElementAt(int.Parse(c.CellValue.InnerText));
                            //                        cellValue = ssi.Text.Text;
                            //                        htSharedTexts.Add(c.CellValue.InnerText, ssi.Text.Text);
                            //                    }
                            //                } else {
                            //                    if (c.CellValue.InnerText != null) {

                            //                        int numberFormatToApply = -1;
                            //                        if (!String.IsNullOrEmpty(strCellFormat)) {
                            //                            numberFormatToApply = Convert.ToInt32(cellFormats[Convert.ToInt16(strCellFormat)].NumberFormatId.Value);
                            //                            cellValue = GetCellFormatValue(c.CellValue.InnerText, numberFormatToApply);
                            //                        } else {

                            //                            cellValue = c.CellValue.InnerText;
                            //                        }
                            //                    } else {
                            //                        cellValue = null;
                            //                    }
                            //                }
                            //            }

                            //            retValues.Add(cellValue);
                            //        }


                            //    }
                            //} while (readerXml.ReadNextSibling());



                            //while (iLastCol < m_ColIndexEnd) {
                            //    retValues.Add(null);
                            //    iLastCol++;
                            //}

                            

                        }
                    }

                }

                while (iLastRow < m_RowIndexEnd) {
                    for (int j = m_ColIndexStart; j <= m_ColIndexEnd; j++) {
                        retValues.Add(null);
                    }
                    iLastRow++;
                }

                if (retValues.Count == 1) {
                    return retValues[0];
                }

                return retValues;
            } catch (Exception ex) {
                throw ex;
            } finally {
                readerXml.Close();
            }
        }

        private void GetRangeRowValues(
            DocumentFormat.OpenXml.OpenXmlReader readerXml, 
            string strRowIndex, 
            ref Hashtable htSharedTexts, 
            CellFormat[] cellFormats,
            ref List<object> retValues) {
            
            int iLastCol = m_ColIndexStart - 1;
            readerXml.ReadFirstChild();
            do {

                if (readerXml.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.Cell)) {
                    bool isColInRange = false;
                    int iCurrColIndex = -1;
                    string strCellFormat = null;
                    if (readerXml.HasAttributes) {
                        string cellRef = readerXml.Attributes.First(a => a.LocalName == "r").Value;
                        var xmlCellFormat = readerXml.Attributes.SingleOrDefault(attr => attr.LocalName == "s");
                        if (xmlCellFormat != null) {
                            strCellFormat = xmlCellFormat.Value;
                        }
                        string strCol = cellRef.Replace(strRowIndex, "");

                        iCurrColIndex = Worksheet.GetColIndexFromLetter(strCol);
                        if (m_ColIndexEnd < iCurrColIndex) {
                            break;
                        }
                        isColInRange = (m_ColIndexStart <= iCurrColIndex);
                    }

                    if (isColInRange) {
                        for (int i = iLastCol + 1; i < iCurrColIndex; i++) {
                            retValues.Add(null);
                        }
                        iLastCol = iCurrColIndex;

                        DocumentFormat.OpenXml.Spreadsheet.Cell c = (DocumentFormat.OpenXml.Spreadsheet.Cell)readerXml.LoadCurrentElement();

                        object cellValue = null;

                        if (c.CellValue != null) {
                            if (c.DataType != null && c.DataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) {
                                if (htSharedTexts.ContainsKey(c.CellValue.InnerText)) {
                                    cellValue = htSharedTexts[c.CellValue.InnerText].ToString();
                                } else {
                                    DocumentFormat.OpenXml.Spreadsheet.SharedStringItem ssi = CurrWorkbookPart.SharedStringTablePart.SharedStringTable.Elements<DocumentFormat.OpenXml.Spreadsheet.SharedStringItem>().ElementAt(int.Parse(c.CellValue.InnerText));
                                    if (ssi != null && ssi.Text != null) {
                                        cellValue = ssi.Text.Text;
                                        htSharedTexts.Add(c.CellValue.InnerText, ssi.Text.Text);
                                    } else if(ssi.InnerText != null) {
                                        cellValue = ssi.InnerText;
                                        htSharedTexts.Add(c.CellValue.InnerText, ssi.InnerText);
                                    } else {
                                        cellValue = null;
                                        htSharedTexts.Add(c.CellValue.InnerText, null);
                                    }
                                }
                            } else {
                                if (c.CellValue.InnerText != null) {

                                    int numberFormatToApply = -1;
                                    if (!String.IsNullOrEmpty(strCellFormat) && cellFormats != null) {
                                        numberFormatToApply = Convert.ToInt32(cellFormats[Convert.ToInt16(strCellFormat)].NumberFormatId.Value);
                                        cellValue = GetCellFormatValue(c.CellValue.InnerText, numberFormatToApply);
                                    } else {

                                        cellValue = c.CellValue.InnerText;
                                    }
                                } else {
                                    cellValue = null;
                                }
                            }
                        }

                        retValues.Add(cellValue);

                        if (iCurrColIndex == m_ColIndexEnd) {
                            break;
                        }
                    }


                }
            } while (readerXml.ReadNextSibling());



            while (iLastCol < m_ColIndexEnd) {
                retValues.Add(null);
                iLastCol++;
            }
        }

        //private List<object> GetRangeValues() {
        //    //List<object> retValues = new List<object>();
        //    //for (int i = 0; i < m_XmlCells.Count; i++) {
        //    //    retValues.Add(Excel.GetCellValue(m_XmlCells[0], CurrSpreadsheetDocument.WorkbookPart));
        //    //}

        //    //return retValues;

        //    int iRowCount = m_RowIndexEnd - m_RowIndexStart;
        //    if (iRowCount < 0) {
        //        int iTmp = m_RowIndexEnd;
        //        m_RowIndexEnd = m_RowIndexStart;
        //        m_RowIndexStart = iTmp;
        //        //iRowCount *= -1;
        //    }
        //    //iRowCount++;

        //    int iColCount = m_ColIndexEnd - m_ColIndexStart;
        //    if (iColCount < 0) {
        //        int iTmp = m_ColIndexEnd;
        //        m_ColIndexEnd = m_ColIndexStart;
        //        m_ColIndexStart = iTmp;
        //        //iColCount *= -1;
        //    }
        //    //iColCount++;

        //    DocumentFormat.OpenXml.Packaging.WorksheetPart wsPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)CurrWorkbookPart.GetPartById(m_Worksheet.XmlSheet.Id);
        //    DocumentFormat.OpenXml.OpenXmlReader readerXml = DocumentFormat.OpenXml.OpenXmlReader.Create(wsPart);
        //    string text;
        //    while (readerXml.Read()) {
        //        if (readerXml.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.Row)) {
        //            //do {
        //            //    if (readerXml.HasAttributes) {
        //            //        string rowNum = readerXml.Attributes.First(a => a.LocalName == "r").Value;
        //            //        Console.Write("rowNum: " + rowNum);
        //            //    }

        //            //} while (readerXml.ReadNextSibling()); // Skip to the next row

        //            if (readerXml.HasAttributes) {
        //                string rowNum = readerXml.Attributes.First(a => a.LocalName == "r").Value;
        //                //Console.Write("rowNum: " + rowNum);
        //            }

        //            readerXml.ReadFirstChild();

        //            do {
        //                if (readerXml.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.Cell)) {
        //                    if (readerXml.HasAttributes) {
        //                        string cellRef = readerXml.Attributes.First(a => a.LocalName == "r").Value;
        //                        //Console.Write("rowNum: " + cellRef);
        //                    }

        //                    DocumentFormat.OpenXml.Spreadsheet.Cell c = (DocumentFormat.OpenXml.Spreadsheet.Cell)readerXml.LoadCurrentElement();

        //                    string cellValue;

        //                    if (c.DataType != null && c.DataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) {
        //                        DocumentFormat.OpenXml.Spreadsheet.SharedStringItem ssi = CurrWorkbookPart.SharedStringTablePart.SharedStringTable.Elements<DocumentFormat.OpenXml.Spreadsheet.SharedStringItem>().ElementAt(int.Parse(c.CellValue.InnerText));

        //                        cellValue = ssi.Text.Text;
        //                    } else {
        //                        cellValue = c.CellValue.InnerText;
        //                    }

        //                    //Console.Out.Write("{0}: {1} ", c.CellReference, cellValue);
        //                }
        //            } while (readerXml.ReadNextSibling());

        //        }
        //        //if (readerXml.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.Cell)) {
        //        //    //DocumentFormat.OpenXml.Spreadsheet.Cell cell = (DocumentFormat.OpenXml.Spreadsheet.Cell)readerXml.ElementType;
        //        //    text = readerXml.GetText();
        //        //    Console.Write(text + " ");
        //        //    //readerXml.ReadFirstChild();
        //        //    //while (readerXml.ReadNextSibling()) {
        //        //    //    readerXml.GetText();
        //        //    //}
        //        //}
        //        //if (readerXml.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.CellValue)) {
        //        //    //DocumentFormat.OpenXml.Spreadsheet.Cell cell = (DocumentFormat.OpenXml.Spreadsheet.Cell)readerXml.ElementType;
        //        //    text = readerXml.GetText();
        //        //    Console.Write(text + " ");
        //        //}
        //    }

        //    return null;
        //}

        //private object GetRangeText() {
        //    //if (m_XmlCells == null || m_XmlCells.Count == 0) {
        //    //    return null;
        //    //}

        //    //if (m_XmlCells.Count == 1) {
        //    //    return Excel.GetCellValue(m_XmlCells[0], CurrSpreadsheetDocument.WorkbookPart);
        //    //}

        //    //object[] retValues = new object[m_Cells.Count];
        //    //for (int i = 0; i < m_Cells.Count; i++) {
        //    //    retValues[i] = Excel.GetCellValue(m_XmlCells[i], CurrSpreadsheetDocument.WorkbookPart);
        //    //}

        //    //return retValues;

        //    return null;
        //}

        //private List<Cell> GetCells() {
        //    int iRowCount = m_RowIndexEnd - m_RowIndexStart;
        //    if (iRowCount < 0) {
        //        int iTmp = m_RowIndexEnd;
        //        m_RowIndexEnd = m_RowIndexStart;
        //        m_RowIndexStart = iTmp;
        //        iRowCount *= -1;
        //    }

        //    int iColCount = m_ColIndexEnd - m_ColIndexStart;
        //    if (iColCount < 0) {
        //        int iTmp = m_ColIndexEnd;
        //        m_ColIndexEnd = m_ColIndexStart;
        //        m_ColIndexStart = iTmp;
        //        iColCount *= -1;
        //    }

        //    List<Cell> cells = new List<Cell>();

        //    return cells;
        //}

        //private List<DocumentFormat.OpenXml.Spreadsheet.Cell> GetXmlCells() {
        //    List<DocumentFormat.OpenXml.Spreadsheet.Cell> cells = new List<DocumentFormat.OpenXml.Spreadsheet.Cell>();

        //    //int iRowCount = m_RowIndexEnd - m_RowIndexStart;
        //    //if (iRowCount < 0) {
        //    //    int iTmp = m_RowIndexEnd;
        //    //    m_RowIndexEnd = m_RowIndexStart;
        //    //    m_RowIndexStart = iTmp;
        //    //    //iRowCount *= -1;
        //    //}
        //    ////iRowCount++;

        //    //int iColCount = m_ColIndexEnd - m_ColIndexStart;
        //    //if (iColCount < 0) {
        //    //    int iTmp = m_ColIndexEnd;
        //    //    m_ColIndexEnd = m_ColIndexStart;
        //    //    m_ColIndexStart = iTmp;
        //    //    //iColCount *= -1;
        //    //}
        //    ////iColCount++;

        //    //DocumentFormat.OpenXml.Packaging.WorksheetPart wsPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)CurrWorkbookPart.GetPartById(m_Worksheet.XmlSheet.Id);
        //    //DocumentFormat.OpenXml.OpenXmlReader readerXml = DocumentFormat.OpenXml.OpenXmlReader.Create(wsPart);
        //    //string text;
        //    //while (readerXml.Read()) {
        //    //    if (readerXml.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.Row)) {
        //    //        //do {
        //    //        //    if (readerXml.HasAttributes) {
        //    //        //        string rowNum = readerXml.Attributes.First(a => a.LocalName == "r").Value;
        //    //        //        Console.Write("rowNum: " + rowNum);
        //    //        //    }

        //    //        //} while (readerXml.ReadNextSibling()); // Skip to the next row

        //    //        if (readerXml.HasAttributes) {
        //    //            string rowNum = readerXml.Attributes.First(a => a.LocalName == "r").Value;
        //    //            //Console.Write("rowNum: " + rowNum);
        //    //        }

        //    //        readerXml.ReadFirstChild();

        //    //        do {
        //    //            if (readerXml.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.Cell)) {
        //    //                if (readerXml.HasAttributes) {
        //    //                    string cellRef = readerXml.Attributes.First(a => a.LocalName == "r").Value;
        //    //                    //Console.Write("rowNum: " + cellRef);
        //    //                }

        //    //                DocumentFormat.OpenXml.Spreadsheet.Cell c = (DocumentFormat.OpenXml.Spreadsheet.Cell)readerXml.LoadCurrentElement();

        //    //                string cellValue;

        //    //                if (c.DataType != null && c.DataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) {
        //    //                    DocumentFormat.OpenXml.Spreadsheet.SharedStringItem ssi = CurrWorkbookPart.SharedStringTablePart.SharedStringTable.Elements<DocumentFormat.OpenXml.Spreadsheet.SharedStringItem>().ElementAt(int.Parse(c.CellValue.InnerText));

        //    //                    cellValue = ssi.Text.Text;
        //    //                } else {
        //    //                    cellValue = c.CellValue.InnerText;
        //    //                }

        //    //                //Console.Out.Write("{0}: {1} ", c.CellReference, cellValue);
        //    //            }
        //    //        } while (readerXml.ReadNextSibling());

        //    //    }
        //    //    //if (readerXml.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.Cell)) {
        //    //    //    //DocumentFormat.OpenXml.Spreadsheet.Cell cell = (DocumentFormat.OpenXml.Spreadsheet.Cell)readerXml.ElementType;
        //    //    //    text = readerXml.GetText();
        //    //    //    Console.Write(text + " ");
        //    //    //    //readerXml.ReadFirstChild();
        //    //    //    //while (readerXml.ReadNextSibling()) {
        //    //    //    //    readerXml.GetText();
        //    //    //    //}
        //    //    //}
        //    //    //if (readerXml.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.CellValue)) {
        //    //    //    //DocumentFormat.OpenXml.Spreadsheet.Cell cell = (DocumentFormat.OpenXml.Spreadsheet.Cell)readerXml.ElementType;
        //    //    //    text = readerXml.GetText();
        //    //    //    Console.Write(text + " ");
        //    //    //}
        //    //}

        //    ////WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
        //    //DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart = CurrWorkbookPart.WorksheetParts.First();


        //    //DocumentFormat.OpenXml.OpenXmlReader reader = DocumentFormat.OpenXml.OpenXmlReader.Create(worksheetPart);
        //    //string text1;
        //    //while (reader.Read()) {
        //    //    if (reader.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.CellValue)) {
        //    //        text1 = reader.GetText();
        //    //        Console.Write(text1 + " ");
        //    //    }
        //    //    //if (reader.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.CellValue)) {
        //    //    //    text = reader.GetText();
        //    //    //    Console.Write(text + " ");
        //    //    //}
        //    //}



        //    //List<DocumentFormat.OpenXml.Spreadsheet.Cell> cells = new List<DocumentFormat.OpenXml.Spreadsheet.Cell>();
        //    ////single row
        //    //if (m_RowIndexStart == m_RowIndexEnd) {
        //    //    var xmlRow = m_Worksheet.GetRow(m_RowIndexStart);
        //    //    GetCellsFromRow(cells, xmlRow);
        //    //} else {
        //    //    UInt32 uiRowIndexStart = Convert.ToUInt32(m_RowIndexStart);
        //    //    UInt32 uiRowIndexEnd = Convert.ToUInt32(m_RowIndexEnd);

        //    //    //DocumentFormat.OpenXml.Packaging.WorksheetPart wsPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)CurrWorkbookPart.GetPartById(m_Worksheet.XmlSheet.Id);
        //    //    //var xmlWsChildren = wsPart.Worksheet.ChildElements.ElementAt(0);
        //    //    //do {
        //    //    //    if (xmlWsChildren is DocumentFormat.OpenXml.Spreadsheet.SheetData) {
        //    //    //        break;
        //    //    //    }
        //    //    //} while (xmlWsChildren != null);

        //    //    //var xmlRowFirst = xmlWsChildren.FirstOrDefault();

        //    //    //var xmlRoot = CurrWorkbookPart.RootElement;
        //    //    //var xmlChild = xmlRoot.ChildElements.FirstOrDefault();
        //    //    //do {
        //    //    //    if (xmlChild is DocumentFormat.OpenXml.Spreadsheet.SheetData) {
        //    //    //        break;
        //    //    //    }             
        //    //    //} while (xmlChild != null);

        //    //    //var xmlRowFirst = xmlChild.FirstOrDefault();

        //    //    //many rows
        //    //    var xmlRowFirst = m_Worksheet.SheetData.First();

        //    //    //var rows = m_Worksheet.SheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>();
        //    //    //for (int i = 0; i < rows.Count(); i++) {
        //    //    //}

        //    //    UInt32 uiLastRowLoaded = 0;
        //    //    //var xmlRowFirst = m_Worksheet.GetRow(m_RowIndexStart);
        //    //    if (xmlRowFirst != null) {
        //    //        DocumentFormat.OpenXml.OpenXmlElement xmlRow = xmlRowFirst;
        //    //        do {

        //    //            UInt32 iRowIndex = (((DocumentFormat.OpenXml.Spreadsheet.Row)xmlRow).RowIndex);
        //    //            if (uiRowIndexEnd < iRowIndex) {
        //    //                break;
        //    //            }

        //    //            if (uiRowIndexStart <= iRowIndex) {
        //    //                while ((uiLastRowLoaded + 1) < iRowIndex) {
        //    //                    GetCellsFromRow(cells, null);
        //    //                    uiLastRowLoaded++;
        //    //                }
        //    //                uiLastRowLoaded = iRowIndex;
        //    //                GetCellsFromRow(cells, (DocumentFormat.OpenXml.Spreadsheet.Row)xmlRow);
        //    //            }

        //    //            xmlRow = xmlRow.NextSibling();
        //    //            //GetCellsFromRow(cells, (DocumentFormat.OpenXml.Spreadsheet.Row)xmlRow);
        //    //        } while (xmlRow != null);
        //    //    }


        //    //}

        //    return cells;
        //}


        //private void GetCellsFromRow(List<DocumentFormat.OpenXml.Spreadsheet.Cell> cells, DocumentFormat.OpenXml.Spreadsheet.Row xmlRow) {
        //    if (xmlRow == null) {
        //        for (int i = m_ColIndexStart; i <= m_ColIndexEnd; i++) {
        //            cells.Add(new DocumentFormat.OpenXml.Spreadsheet.Cell());
        //        }
        //    } else {
        //        UInt32 uiColIndexStart = Convert.ToUInt32(m_ColIndexStart);
        //        UInt32 uiColIndexEnd = Convert.ToUInt32(m_ColIndexEnd);

        //        var xmlCellFirst = xmlRow.First();

        //        if (xmlCellFirst != null) {
        //            DocumentFormat.OpenXml.OpenXmlElement xmlCell = xmlCellFirst;
        //            do {

        //                string cellReference = (((DocumentFormat.OpenXml.Spreadsheet.Cell)xmlCell).CellReference);
        //                int iRow = -1;
        //                string strCol = null;
        //                Worksheet.SeparateCellReference(cellReference, out iRow, out strCol);
        //                UInt32 uiColIndex = Convert.ToUInt32(Worksheet.GetColIndexFromLetter(strCol));
        //                if (uiColIndexEnd < uiColIndex) {
        //                    break;
        //                }

        //                if (uiColIndexStart <= uiColIndex) {
        //                    cells.Add((DocumentFormat.OpenXml.Spreadsheet.Cell)xmlCell);
        //                }

        //                xmlCell = xmlCell.NextSibling();
        //                //GetCellsFromRow(cells, (DocumentFormat.OpenXml.Spreadsheet.Row)xmlRow);
        //            } while (xmlCell != null);
        //        }

        //        ////single cell
        //        //var xmlCell = m_Worksheet.GetCell(xmlRow, m_ColIndexStart);
        //        //if (xmlCell == null) {
        //        //    xmlCell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
        //        //}
        //        //cells.Add(xmlCell);
        //    } 
        //}
        #endregion
    }
}
