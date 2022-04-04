using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kamsyk.ExcelOpenXml.ExcelActiveX {
    public abstract class BaseExcelObject {
        private static SpreadsheetDocument m_SpreadsheetDocument = null;
        protected SpreadsheetDocument CurrSpreadsheetDocument {
            get { return m_SpreadsheetDocument; }
            set { m_SpreadsheetDocument = value; }
        }

        protected WorkbookPart CurrWorkbookPart {
            get { return (WorkbookPart)m_SpreadsheetDocument.WorkbookPart; }
        }

        //protected static DocumentFormat.OpenXml.Spreadsheet.Row GetRow(DocumentFormat.OpenXml.Spreadsheet.SheetData sheetData, int iRowIndex) {
        //    if (sheetData.ChildElements == null || sheetData.ChildElements.Count == 0) {
        //        return null;
        //    }

        //    //var row = m_Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>().Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().ElementAt(iRowIndex);
        //    //var row = m_Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>().Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Where(r => r.RowIndex == iRowIndex).FirstOrDefault();

        //    var row = sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Where(r => r.RowIndex == iRowIndex).FirstOrDefault();

        //    return row;
        //}

        //protected static DocumentFormat.OpenXml.Spreadsheet.SheetData GetSheetData(DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet) {
        //    foreach (var child in worksheet.ChildElements) {
        //        if (child is DocumentFormat.OpenXml.Spreadsheet.SheetData) {
        //            return (DocumentFormat.OpenXml.Spreadsheet.SheetData)child;
        //        }
        //    }

        //    return null;
        //}
    }
}
