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
    public class RangeTests : BaseTests {
        #region TestMethods
        [TestMethod()]
        public void SetSingleRange_GetRangeValueString_IsString() {
            
            string excelFileName = null;
            try {
                // Arrange
                string strValue = "TextValueAbcd";
                string strRangeAddress = GetRandomRangeAddress();
                excelFileName = PreapreExcelValue(strRangeAddress, strValue);

                // Act
                Application excelApp = new Application();
                var xmbWb = excelApp.Workbooks.Open(excelFileName);
                var xmlWs = xmbWb.Worksheets[0];
                Range r = xmlWs.GetRange(strRangeAddress);
                object oResult = r.Value;
                
                // Assert
                Assert.IsTrue(oResult != null && (oResult.GetType() == typeof(String)) && !String.IsNullOrEmpty(oResult.ToString()) && oResult.ToString() == strValue);
            } catch (Exception ex) {
                throw ex;
            } finally {
                try {
                    File.Delete(excelFileName);
                } catch { }
            }

        }

        [TestMethod()]
        public void SetSingleRange_GetRangeValueString_IsNumber() {
                        
            string excelFileName = null;
            try {
                // Arrange
                double dValue = new Random().Next(1, 100000);
                string strRangeAddress = GetRandomRangeAddress();
                string numberFormat = "0.00";
                excelFileName = PreapreExcelValue(strRangeAddress, dValue, numberFormat);

                
                // Act
                Application excelApp = new Application();
                var xmbWb = excelApp.Workbooks.Open(excelFileName);
                var xmlWs = xmbWb.Worksheets[0];
                DateTime dtStart = DateTime.Now;
                Range r = xmlWs.GetRange(strRangeAddress);
                object oResult = r.Value;
                DateTime dtStop = DateTime.Now;

                TimeSpan ts = dtStop.Subtract(dtStart);
                var ms = ts.Milliseconds;

                // Assert
                Assert.IsTrue(oResult != null && (oResult.GetType() == typeof(double)) && (double)oResult == dValue);
            } catch (Exception ex) {
                throw ex;
            } finally {
                try {
                    File.Delete(excelFileName);
                } catch { }
            }

        }

        [TestMethod()]
        public void SetSingleRange_GetRangeValueString_IsDateTime() {

            string excelFileName = null;
            try {
                // Arrange
                DateTime dValue = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
                string strRangeAddress = GetRandomRangeAddress();
                string numberFormat = "m/d/yyyy";
                excelFileName = PreapreExcelValue(strRangeAddress, dValue, numberFormat);


                // Act
                Application excelApp = new Application();
                var xmbWb = excelApp.Workbooks.Open(excelFileName);
                var xmlWs = xmbWb.Worksheets[0];
                Range r = xmlWs.GetRange(strRangeAddress);
                object oResult = r.Value;

                // Assert
                Assert.IsTrue(oResult != null && (oResult.GetType() == typeof(DateTime)) && (DateTime)oResult == dValue);
            } catch (Exception ex) {
                throw ex;
            } finally {
                try {
                    File.Delete(excelFileName);
                } catch { }
            }

        }

        [TestMethod()]
        public void SetRow_GetRangeValueString_IsString() {
            string excelFileName = null;
            try {
                // Arrange
                string strValue = "StringValueTest_sdfjh sdjahfgsdjga afsdkjh gfsdkjh gfsdkjhsa";
                string strRangeAddress = "3:3";
                excelFileName = PreapreExcelValue(strRangeAddress, strValue);


                // Act
                Application excelApp = new Application();
                var xmbWb = excelApp.Workbooks.Open(excelFileName);
                var xmlWs = xmbWb.Worksheets[0];
                Range r = xmlWs.GetRange(strRangeAddress);
                object oResult = r.Value;

                // Assert
                bool isOK = true;
                List<object> retValues = (List<object>)oResult;
                isOK = isOK && retValues.Count == Worksheet.GetColIndexFromLetter(Worksheet.MAX_RANGE_COLUMN);
                if (isOK) {
                    foreach (var oItem in retValues) {
                        if (oItem.ToString() != strValue) {
                            isOK = false;
                            break;
                        }
                    }
                }
                                
                Assert.IsTrue(isOK);
            } catch (Exception ex) {
                throw ex;
            } finally {
                try {
                    File.Delete(excelFileName);
                } catch { }
            }
        }

        [TestMethod()]
        public void GetRangeValueString_FromFileSharedStringNull_IsStringValueOk() {
            string excelFileName = null;
            try {
                // Arrange
                string assFile = System.Reflection.Assembly.GetExecutingAssembly().Location;
                FileInfo fi = new FileInfo(assFile);
                DirectoryInfo di = new DirectoryInfo(fi.Directory.FullName);
                string excFolder = di.Parent.Parent.FullName;
                excFolder = Path.Combine(excFolder, "TestFiles");
                excelFileName = Path.Combine(excFolder, "SrbRegetSuppliers.xlsx");
                string strRangeAddress = "H256";
                

                // Act
                Application excelApp = new Application();
                var xmbWb = excelApp.Workbooks.Open(excelFileName);
                var xmlWs = xmbWb.Worksheets[0];
                Range r = xmlWs.GetRange(strRangeAddress);
                object oResult = r.Value;

                // Assert
                Assert.IsTrue(oResult.ToString() == "+39 0521 3111");
            } catch (Exception ex) {
                throw ex;
            } 
        }

        [TestMethod()]
        public void SetColumn_GetRangeValueString_IsString() {
            string excelFileName = null;
            try {
                // Arrange
                string strValue = "StringValueTest_sdfjh sdjahfgsdjga afsdkjh gfsdkjh gfsdkjhsa";
                string strRangeAddress = "B:B";
                excelFileName = PreapreExcelValue(strRangeAddress, strValue);


                // Act
                Application excelApp = new Application();
                var xmbWb = excelApp.Workbooks.Open(excelFileName);
                var xmlWs = xmbWb.Worksheets[0];
                Range r = xmlWs.GetRange(strRangeAddress);
                object oResult = r.Value;

                // Assert
                bool isOK = true;
                List<object> retValues = (List<object>)oResult;
                isOK = isOK && retValues.Count == Worksheet.MAX_RANGE_ROW_INDEX;
                if (isOK) {
                    foreach (var oItem in retValues) {
                        if (oItem == null) {
                            isOK = false;
                            break;
                        }
                        if (oItem.ToString() != strValue) {
                            isOK = false;
                            break;
                        }
                    }
                }

                Assert.IsTrue(isOK);
            } catch (Exception ex) {
                throw ex;
            } finally {
                try {
                    File.Delete(excelFileName);
                } catch { }
            }
        }

        [TestMethod()]
        public void SetRange_GetCoupleRangeContainsAllValueString_IsOk() {
            string excelFileName = null;

            try {
                // Arrange
                string strValue = "StringValueTest_sdfjh sdjahfgsdjga afsdkjh gfsdkjh gfsdkjhsa";
                string strRangeAddress = "B1:C2";
                excelFileName = PreapreExcelValue(strRangeAddress, strValue);

                // Act
                Application excelApp = new Application();
                var xmbWb = excelApp.Workbooks.Open(excelFileName);
                var xmlWs = xmbWb.Worksheets[0];
                Range r = xmlWs.GetRange("A1:D100");
                object oResult = r.Value;

                // Assert
                bool isOK = true;
                List<object> retValues = (List<object>)oResult;
                isOK = isOK && retValues.Count == 400;
                if (isOK) {
                    int stringCellsCount = 0;
                    foreach (var oItem in retValues) {
                        if (oItem != null && oItem.ToString() == strValue) {
                            stringCellsCount++;
                            
                        }
                    }

                    if (stringCellsCount != 4) {
                        isOK = false;
                    }
                }

                Assert.IsTrue(isOK);
            } catch (Exception ex) {
                throw ex;
            } finally {
                try {
                    File.Delete(excelFileName);
                } catch { }
            }
        }

        [TestMethod()]
        public void SetRange_GetCoupleRangeContainsNoneValueString_IsOk() {
            string excelFileName = null;

            try {
                // Arrange
                string strValue = "StringValueTest_sdfjh sdjahfgsdjga afsdkjh gfsdkjh gfsdkjhsa";
                string strRangeAddress = "B1:C2";
                excelFileName = PreapreExcelValue(strRangeAddress, strValue);

                // Act
                Application excelApp = new Application();
                var xmbWb = excelApp.Workbooks.Open(excelFileName);
                var xmlWs = xmbWb.Worksheets[0];
                Range r = xmlWs.GetRange("F1:G10");
                object oResult = r.Value;

                // Assert
                bool isOK = true;
                List<object> retValues = (List<object>)oResult;
                isOK = isOK && retValues.Count == 20;
                if (isOK) {
                    int stringCellsCount = 0;
                    foreach (var oItem in retValues) {
                        if (oItem != null && oItem.ToString() == strValue) {
                            stringCellsCount++;

                        }
                    }

                    if (stringCellsCount != 0) {
                        isOK = false;
                    }
                }

                Assert.IsTrue(isOK);
            } catch (Exception ex) {
                throw ex;
            } finally {
                try {
                    File.Delete(excelFileName);
                } catch { }
            }
        }


        [TestMethod()]
        public void SetRange_GetBiggerRangeSeveralCellsTime_IsOk() {
            string excelFileName = null;

            try {
                // Arrange
                string strValue = "StringValueTest_sdfjh sdjahfgsdjga afsdkjh gfsdkjh gfsdkjhsa";
                string strRangeAddress = "A1:J450";
                excelFileName = PreapreExcelValue(strRangeAddress, strValue);

                // Act
                int ms = 0;
                Application excelApp = new Application();
                var xmbWb = excelApp.Workbooks.Open(excelFileName);
                var xmlWs = xmbWb.Worksheets[0];
                DateTime dtStart = new DateTime();
                using (StreamWriter sw = new StreamWriter(@"c:\temp\exctest.txt")) {
                    for (int i = 0; i < 400; i++) {
                        string strRandomRangeAddress = GetRandomRangeAddress();
                        Range r = xmlWs.GetRange(strRandomRangeAddress);
                        var oResult = r.Value;
                        if (oResult == null) {
                            sw.WriteLine(strRandomRangeAddress + ":null");
                        } else {
                            sw.WriteLine(strRandomRangeAddress +":" + oResult.ToString());
                        }
                    }

                    DateTime dtStop = new DateTime();

                    // Assert
                    TimeSpan ts = dtStop.Subtract(dtStart);
                    ms = ts.Seconds;

                    sw.WriteLine(ms.ToString());
                }


                Assert.IsTrue(ms < 50);
            } catch (Exception ex) {
                throw ex;
            } finally {
                try {
                    File.Delete(excelFileName);
                } catch { }
            }
        }

        [TestMethod()]
        public void SetRange_GetCoupleRangeContainsPartialValueString_IsOk() {
            string excelFileName = null;

            try {
                // Arrange
                string strValue = "StringValueTest_sdfjh sdjahfgsdjga afsdkjh gfsdkjh gfsdkjhsa";
                string strRangeAddress = "B1:C2";
                excelFileName = PreapreExcelValue(strRangeAddress, strValue);

                // Act
                Application excelApp = new Application();
                var xmbWb = excelApp.Workbooks.Open(excelFileName);
                var xmlWs = xmbWb.Worksheets[0];
                Range r = xmlWs.GetRange("C1:D10");
                object oResult = r.Value;

                // Assert
                bool isOK = true;
                List<object> retValues = (List<object>)oResult;
                isOK = isOK && retValues.Count == 20;
                if (isOK) {
                    int stringCellsCount = 0;
                    foreach (var oItem in retValues) {
                        if (oItem != null && oItem.ToString() == strValue) {
                            stringCellsCount++;

                        }
                    }

                    if (stringCellsCount != 2) {
                        isOK = false;
                    }
                }

                Assert.IsTrue(isOK);
            } catch (Exception ex) {
                throw ex;
            } finally {
                try {
                    File.Delete(excelFileName);
                } catch { }
            }
        }

        [TestMethod()]
        public void SetBigRange_GetRangeValueString_IsString() {
            string excelFileName = null;
            try {
                // Arrange
                int iRowsCount = 10000;
                string strValue = "StringValueTest_sdfjh sdjahfgsdjga afsdkjh gfsdkjh gfsdkjhsa";
                string strRangeAddress = "A1:Z" + iRowsCount;
                excelFileName = PreapreExcelValue(strRangeAddress, strValue);
                
                // Act
                Application excelApp = new Application();
                var xmbWb = excelApp.Workbooks.Open(excelFileName);
                var xmlWs = xmbWb.Worksheets[0];
                Range r = xmlWs.GetRange(strRangeAddress);
                object oResult = r.Value;

                // Assert
                bool isOK = true;
                List<object> retValues = (List<object>)oResult;
                isOK = isOK && retValues.Count == iRowsCount * Worksheet.GetColIndexFromLetter("Z");
                if (isOK) {
                    foreach (var oItem in retValues) {
                        if (oItem.ToString() != strValue) {
                            isOK = false;
                            break;
                        }
                    }
                }

                Assert.IsTrue(isOK);
            } catch (Exception ex) {
                throw ex;
            } finally {
                try {
                    File.Delete(excelFileName);
                } catch { }
            }
        }


        [TestMethod()]
        public void NotXlsxXlsm_NotOpened() {
            string excelFileName = null;
            try {
                // Arrange
                excelFileName = "ac.xls";

                // Act
                bool isExcOk = false;
                try {
                    Application excelApp = new Application();
                    var xmbWb = excelApp.Workbooks.Open(excelFileName);
                } catch (Exception ex) {
                    isExcOk = (ex.Message == "Only .xlsx, .xltx, .xlsm files are supported");
                }

                // Assert
                Assert.IsTrue(isExcOk);

            } catch (Exception ex) {
                throw ex;
            } finally {
                try {
                    File.Delete(excelFileName);
                } catch { }
            }
        }

        #endregion

        #region Methods
        //private string GetTmpFolder() {
        //    string strFileLoc = System.Reflection.Assembly.GetExecutingAssembly().Location;
        //    FileInfo fi = new FileInfo(strFileLoc);
        //    string strFolder = fi.DirectoryName;
        //    strFolder = Path.Combine(strFolder, "Excel");

        //    if (Directory.Exists(strFolder)) {
        //        foreach (string file in Directory.GetFiles(strFolder)) {
        //            try {
        //                File.Delete(file);
        //            } catch { }
        //        }
        //    } else {
        //        Directory.CreateDirectory(strFolder);
        //    }

        //    return strFolder;
        //}

        //private string GetExcelFileName() {
        //    string strPureName = "TmpExcelTest";
        //    string strFilePath = Path.Combine(GetTmpFolder(), strPureName + ".xlsx");
        //    int iIndex = 1;
        //    while (File.Exists(strFilePath)) {
        //        strFilePath = Path.Combine(GetTmpFolder(), strPureName + "_" + iIndex + ".xlsx");
        //        iIndex++;
        //    }

        //    //string strFilePath = Excel.GetOrigFileName(GetTmpFolder(), strPureName, ".xlsx");

        //    return strFilePath;
        //}

        private string GetRandomRangeAddress() {
            Random rnd = new Random();
            int iRow = rnd.Next(100);
            int iCol = rnd.Next(65, 90);
            char character = (char)iCol;
            string strCol = character.ToString();

            int iTmp = rnd.Next(10);
            if (iTmp > 5) {
                //2 letters in culumn AA, CG ...
                iCol = rnd.Next(65, 90);
                character = (char)iCol;
                strCol += character.ToString();
            } 

            string rangeAddress = strCol + iRow;

            return rangeAddress;
           
        }

        //private string PreapreExcelValue(string strRangeAddress, object oValue) {
        //    return PreapreExcelValue(strRangeAddress, oValue, null);
        //}

        //private string  PreapreExcelValue(string strRangeAddress, object oValue, string numberFormat) {
        //    Microsoft.Office.Interop.Excel.Application exApp = new Microsoft.Office.Interop.Excel.Application();
        //    exApp.Visible = true;
        //    var wb = exApp.Workbooks.Add();

        //    var ws = wb.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
        //    var range = ws.Range[strRangeAddress];
        //    if (!String.IsNullOrEmpty(numberFormat)) {
        //        range.NumberFormat = numberFormat;
        //    }
        //    range.Value = oValue;
        //    string excelFileName = GetExcelFileName();
        //    wb.SaveAs(excelFileName);
        //    wb.Close();
        //    exApp.Quit();

        //    return excelFileName;
        //}

        #endregion
    }
}