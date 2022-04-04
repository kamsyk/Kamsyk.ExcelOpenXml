using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kamsyk.ExcelOpenXml.ExcelActiveX {
    public class Workbooks : BaseNotInstantiable{
        #region Propereties
        
        #endregion

        #region Methods
        public Workbook Open(string Filename) {
            FileInfo fi = new FileInfo(Filename);
            if (fi.Extension.Trim().ToLower() != ".xlsx" &&
                fi.Extension.Trim().ToLower() != ".xlsm" &&
                fi.Extension.Trim().ToLower() != ".xlts") {
                throw new Exception("Only .xlsx, .xltx, .xlsm files are supported");
            }

            CurrSpreadsheetDocument = SpreadsheetDocument.Open(Filename, true);

            var workbookPart = CurrSpreadsheetDocument.WorkbookPart;
            Workbook wb = new Workbook();
            wb.SetWorkbookPart(workbookPart);

            return wb;
        }

        public Workbook Add() {
            Excel excel = new Excel();
            CurrSpreadsheetDocument = excel.GetNewXlsDocMemory();

            var workbookPart = excel.AddWorkbook(CurrSpreadsheetDocument);
            Workbook wb = new Workbook();
            wb.SetWorkbookPart(workbookPart);

            return wb;
            
        }
        #endregion
    }
}
