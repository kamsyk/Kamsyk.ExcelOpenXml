using Microsoft.VisualStudio.TestTools.UnitTesting;
using Kamsyk.ExcelOpenXml.ExcelActiveX;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Kamsyk.ExcelOpenXml.ExcelActiveX.Tests {
    [TestClass()]
    public class WorksheetTests : BaseTests {
        [TestMethod()]
        public void GetCellAfterEmtyRowCell_NotNull() {
            string excelFileName = null;
            try {
                // Arrange
                string strValue = "StringValueTest_sdfjh sdjahfgsdjga afsdkjh gfsdkjh gfsdkjhsa";
                string strRangeAddress = "D3";
                excelFileName = PreapreExcelValue(strRangeAddress, strValue);


                // Act
                Application excelApp = new Application();
                var xmbWb = excelApp.Workbooks.Open(excelFileName);
                var xmlWs = xmbWb.Worksheets[0];
                Cell cell = xmlWs.Cells(3, 4);
                object oResult = cell.Value;

                // Assert
                Assert.IsTrue(oResult != null);
            } catch (Exception ex) {
                throw ex;
            } finally {
                try {
                    File.Delete(excelFileName);
                } catch { }
            }
        }
    }
}