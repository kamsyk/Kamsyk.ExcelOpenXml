
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Kamsyk.ExcelOpenXml.ExcelActiveX {
    public class Worksheet : BaseNotInstantiable {
        #region Constants
        public const int MAX_RANGE_ROW_INDEX = 1048576;
        public const string MAX_RANGE_COLUMN = "XFD";
        #endregion

        #region Properties
        private DocumentFormat.OpenXml.Spreadsheet.Sheet m_Sheet = null;
        internal DocumentFormat.OpenXml.Spreadsheet.Sheet XmlSheet {
            get { return m_Sheet; }
        }
        internal DocumentFormat.OpenXml.Packaging.WorksheetPart WorksheetPart {
            get {
                return ((DocumentFormat.OpenXml.Packaging.WorksheetPart)CurrWorkbookPart.GetPartById(m_Sheet.Id));
            }
        }
        private DocumentFormat.OpenXml.Spreadsheet.Worksheet m_Worksheet {
            get { return WorksheetPart.Worksheet; }
        }

        private DocumentFormat.OpenXml.Spreadsheet.SheetData _m_SheetData = null;
        internal DocumentFormat.OpenXml.Spreadsheet.SheetData SheetData {
            get {
                if (_m_SheetData == null) {
                    _m_SheetData = GetSheetData();
                }

                return _m_SheetData;
            }
        }


        #endregion

        #region Methods
        internal void SetSheet(DocumentFormat.OpenXml.Spreadsheet.Sheet sheet) {
            m_Sheet = sheet;
        }

        public Range GetRange(string strRange) {
            int iRow1 = -1;
            string strCol1 = null;
            int iRow2 = -1;
            string strCol2 = null;

            GetCellAddress(
                strRange, 
                out iRow1, 
                out strCol1, 
                out iRow2, 
                out strCol2);

            Range r = null;
            if (iRow1 == -1) {
                //whole column
                iRow1 = 1;
                iRow2 = MAX_RANGE_ROW_INDEX;

                r = new Range(
                    this,
                    iRow1,
                    GetColIndexFromLetter(strCol1),
                    iRow2,
                    GetColIndexFromLetter(strCol2));

                return r;
            }

            if (strCol1 == null) {
                //whole row
                strCol1 = "A";
                strCol2 = "XFD";

                r = new Range(
                    this,
                    iRow1,
                    GetColIndexFromLetter(strCol1),
                    iRow2,
                    GetColIndexFromLetter(strCol2));

                return r;
            }

            r = new Range(
                this,
                iRow1, 
                GetColIndexFromLetter(strCol1),
                iRow2,
                GetColIndexFromLetter(strCol2));

            return r;

            //DocumentFormat.OpenXml.Spreadsheet.Row row = GetRow(iRow1);


            //return null;
        }

        //public Range GetRange(Cell cell1, Cell cell2) {
            
        //    return null;
        //}

        public Cell Cells(int iRow, int iColumn) {
            Cell cell = new Cell(this, iRow, iColumn, iRow, iColumn);

            Cell xmlCell = GetCell(iRow, iColumn);

            
            return cell;
        }

        private Cell GetCell(int iRow, int iCol) {
            //DocumentFormat.OpenXml.Packaging.WorksheetPart wsPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)CurrWorkbookPart.GetPartById(XmlSheet.Id);
            //DocumentFormat.OpenXml.OpenXmlReader readerXml = DocumentFormat.OpenXml.OpenXmlReader.Create(wsPart);

            //while (readerXml.Read()) {
            //    if (readerXml.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.Row)) {
            //        if (readerXml.HasAttributes) {
            //            string strRowIndex = readerXml.Attributes.First(a => a.LocalName == "r").Value;
            //            int iCurrRowIndex = Convert.ToInt32(strRowIndex);
            //            if (iCurrRowIndex == iRow) {
            //                readerXml.ReadFirstChild();
            //                do {

            //                    if (readerXml.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.Cell)) {
            //                        if (readerXml.HasAttributes) {
            //                            string cellRef = readerXml.Attributes.First(a => a.LocalName == "r").Value;
            //                            string strCol = cellRef.Replace(strRowIndex, "");

            //                            int iCurrColIndex = Worksheet.GetColIndexFromLetter(strCol);
            //                            if (iCurrColIndex == iCol) {

            //                            } else if (iCurrColIndex > iCol) {
            //                                return new Cell(this, iRow, iCol);
            //                            }
            //                        }
            //                    }
            //                } while (readerXml.ReadNextSibling());
            //            } else if (iCurrRowIndex > iRow) {
            //                return new Cell(this, iRow, iCol);
            //            }
            //        }
            //    }
            //}

            return new Cell(this, iRow, iCol);
        }

        private DocumentFormat.OpenXml.Spreadsheet.SheetData GetSheetData() {
            //foreach (var child in m_Worksheet.ChildElements) {
            //    if (child is DocumentFormat.OpenXml.Spreadsheet.SheetData) {
            //        return (DocumentFormat.OpenXml.Spreadsheet.SheetData)child;
            //    }
            //}

            //return null;
            return m_Worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.SheetData>().FirstOrDefault();
        }



        private void GetCellAddress(
            string cellAddress, 
            out int iRow1, 
            out string strCol1, 
            out int iRow2, 
            out string strCol2) {

            iRow1 = -1;
            strCol1 = null;

            string[] items = cellAddress.Split(':');
            SeparateCellReference(items[0], out iRow1, out strCol1);

            if (items.Length == 1) {
                iRow2 = iRow1;
                strCol2 = strCol1;
            } else {
                SeparateCellReference(items[1], out iRow2, out strCol2);
            }

            if (String.IsNullOrEmpty(strCol1)) {
                strCol1 = null;
            }

            if (String.IsNullOrEmpty(strCol2)) {
                strCol2 = null;
            }
        }

        internal static void SeparateCellReference(
            string cellAddress, 
            out int iRow, 
            out string strCol) {
            iRow = -1;
            strCol = null;

            //int iPosRowStart = -1;
            //for (int i = 0; i < cellAddress.Length; i++) {
            //    short iNum = -1;
            //    if (Int16.TryParse(cellAddress.Substring(i, 1), out iNum)) {
            //        iPosRowStart = i;
            //        break;
            //    } 
            //}

            //if (iPosRowStart > 0) {
            //    strCol = cellAddress.Substring(0, iPosRowStart);
            //} else {
            //    //whole column - nax rowindex = 1 048 576
            //}

            //if (iPosRowStart > -1) {
            //    iRow = int.Parse(cellAddress.Substring(iPosRowStart));
            //} else {
            //    //whole row - mal col XFD
            //}


            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellAddress);
            strCol = match.Value;

            regex = new Regex("[0-9]+");
            match = regex.Match(cellAddress);

            short iNum = -1;
            if (Int16.TryParse(match.Value, out iNum)) {
                iRow = iNum;
            }
        }

        public static int GetColIndexFromLetter(string columnReference) {
            int iCol = -1;
            int iMulitplier = 1;

            foreach (char c in columnReference.ToCharArray().Reverse()) {
                iCol += iMulitplier * ((int)c - 64);

                iMulitplier = iMulitplier * 26;
            }

            iCol++;

            return iCol;
        }

        private string GetLetterFromColIndex(int colIndex) {
            //int dividend = colIndex;
            //string columnName = String.Empty;
            //int modulo;

            //while (dividend > 0) {
            //    modulo = (dividend - 1) % 26;
            //    columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
            //    dividend = (int)((dividend - modulo) / 26);
            //}

            //return columnName;

            return CommonExcel.GetLetterFromColIndex(colIndex);
        }

        internal DocumentFormat.OpenXml.Spreadsheet.Row GetRow(int iRowIndex) {
            if (SheetData.ChildElements == null || SheetData.ChildElements.Count == 0) {
                return null;
            }

            //var row = m_Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>().Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().ElementAt(iRowIndex);
            //var row = m_Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>().Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Where(r => r.RowIndex == iRowIndex).FirstOrDefault();

            var row = SheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Where(r => r.RowIndex == iRowIndex).FirstOrDefault();

            return row;
        }

        internal DocumentFormat.OpenXml.Spreadsheet.Cell GetCell(DocumentFormat.OpenXml.Spreadsheet.Row xmlRow, int iColIndex) {
            if (SheetData.ChildElements == null || SheetData.ChildElements.Count == 0) {
                return null;
            }

            string cellReference = GetLetterFromColIndex(iColIndex) + xmlRow.RowIndex;
            var cell = xmlRow.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Where(c => c.CellReference == cellReference).FirstOrDefault();

            return cell;
        }
        #endregion
    }
}
