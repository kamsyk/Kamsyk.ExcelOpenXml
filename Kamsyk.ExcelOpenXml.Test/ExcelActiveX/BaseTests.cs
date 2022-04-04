using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kamsyk.ExcelOpenXml.ExcelActiveX.Tests {
    [TestClass()]
    public class BaseTests {
        protected string PreapreExcelValue(string strRangeAddress, object oValue) {
            return PreapreExcelValue(strRangeAddress, oValue, null);
        }

        protected string PreapreExcelValue(string strRangeAddress, object oValue, string numberFormat) {
            Microsoft.Office.Interop.Excel.Application exApp = new Microsoft.Office.Interop.Excel.Application();
            exApp.Visible = true;
            var wb = exApp.Workbooks.Add();

            var ws = wb.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
            var range = ws.Range[strRangeAddress];
            if (!String.IsNullOrEmpty(numberFormat)) {
                range.NumberFormat = numberFormat;
            }
            range.Value = oValue;
            string excelFileName = GetExcelFileName();
            wb.SaveAs(excelFileName);
            wb.Close();
            exApp.Quit();

            return excelFileName;
        }

        protected string GetExcelFileName() {
            string strPureName = "TmpExcelTest";
            string strFilePath = Path.Combine(GetTmpFolder(), strPureName + ".xlsx");
            int iIndex = 1;
            while (File.Exists(strFilePath)) {
                strFilePath = Path.Combine(GetTmpFolder(), strPureName + "_" + iIndex + ".xlsx");
                iIndex++;
            }

            //string strFilePath = Excel.GetOrigFileName(GetTmpFolder(), strPureName, ".xlsx");

            return strFilePath;
        }

        protected string GetTmpFolder() {
            string strFileLoc = System.Reflection.Assembly.GetExecutingAssembly().Location;
            FileInfo fi = new FileInfo(strFileLoc);
            string strFolder = fi.DirectoryName;
            strFolder = Path.Combine(strFolder, "Excel");

            if (Directory.Exists(strFolder)) {
                foreach (string file in Directory.GetFiles(strFolder)) {
                    try {
                        File.Delete(file);
                    } catch { }
                }
            } else {
                Directory.CreateDirectory(strFolder);
            }

            return strFolder;
        }
    }
}
