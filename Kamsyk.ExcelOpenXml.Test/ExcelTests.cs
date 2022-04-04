using Microsoft.VisualStudio.TestTools.UnitTesting;
using Kamsyk.ExcelOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Kamsyk.ExcelOpenXml.Tests {
    [TestClass()]
    public class ExcelTests {
        #region Test Methods
        [TestMethod()]
        public void LockSheetPwd_Save_UnlockwoPwd() {
            string excelFileName = null;
            string unlockFileName = null;
            try {
                // Arrange
                Microsoft.Office.Interop.Excel.Application exApp = new Microsoft.Office.Interop.Excel.Application();
                exApp.Visible = true;
                var axWb = exApp.Workbooks.Add();
                var axWs = axWb.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;

                //set values
                string strRange = "B1:B2";
                var axRange = axWs.Range[strRange];
                axRange.Value = "ewa >wq tqw < t3532 d";

                axWs.Protect("Password");
                excelFileName = GetExcelFileName();
                axWb.SaveAs(excelFileName);
                axWb.Close();
                exApp.Quit();

                // Act
                unlockFileName = Excel.Unlock(excelFileName);

                // Assert
                //Lock
                exApp = new Microsoft.Office.Interop.Excel.Application();
                exApp.Visible = true;
                axWb = exApp.Workbooks.Open(excelFileName);
                axWs = axWb.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
                axRange = axWs.Range[strRange];
                try {
                    axRange.Value = "new value ssdggsd";
                } catch (Exception ex) {
                    if (ex.Message.IndexOf("The cell or chart that you are trying to change is protected and therefore read-only") < 0 &&
                        ex.Message.IndexOf("Pokoušíte ze změnit zamknutou buňku nebo zamknutý graf, které jsou proto jen pro čtení") < 0 &&
                        ex.Message.IndexOf("protected sheet") < 0 &&
                        ex.Message.IndexOf("zamknutém listu") < 0) {
                        Assert.Fail();
                    }
                }
                exApp.DisplayAlerts = false;
                axWb.Close();
                exApp.Quit();

                //Unlock
                exApp = new Microsoft.Office.Interop.Excel.Application();
                exApp.Visible = true;
                axWb = exApp.Workbooks.Open(unlockFileName);
                axWs = axWb.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
                axRange = axWs.Range[strRange];
                try {
                    axRange.Value = "new value ssdggsd";
                } catch {
                    Assert.Fail();
                }
                exApp.DisplayAlerts = false;
                axWb.Close();
                exApp.Quit();

                Assert.IsTrue(1 == 1);

            } catch (Exception ex) {
                throw ex;
            } finally {
                try {
                    File.Delete(excelFileName);
                    File.Delete(unlockFileName);
                } catch { }
            }
        }

        [TestMethod()]
        public void LockLargeSheetPwd_Save_UnlockwoPwd() {
            string excelFileName = null;
            string unlockFileName = null;
            try {
                // Arrange
                Microsoft.Office.Interop.Excel.Application exApp = new Microsoft.Office.Interop.Excel.Application();
                exApp.Visible = true;
                var axWb = exApp.Workbooks.Add();
                var axWs = axWb.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;

                //set values
                string strRange = "B:B";
                var axRange = axWs.Range[strRange];
                axRange.Value = "ewa >wq tqw < t3532 d";

                axWs.Protect("Password");
                excelFileName = GetExcelFileName();
                axWb.SaveAs(excelFileName);
                axWb.Close();
                exApp.Quit();

                // Act
                unlockFileName = Excel.Unlock(excelFileName);

                // Assert
                exApp = new Microsoft.Office.Interop.Excel.Application();
                exApp.Visible = true;
                axWb = exApp.Workbooks.Open(excelFileName);
                axWs = axWb.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
                axRange = axWs.Range[strRange];
                try {
                    axRange.Value = "new value ssdggsd";
                } catch (Exception ex) {
                    if (ex.Message.IndexOf("The cell or chart that you are trying to change is protected and therefore read-only") < 0 &&
                        ex.Message.IndexOf("Pokoušíte ze změnit zamknutou buňku nebo zamknutý graf, které jsou proto jen pro čtení") < 0 &&
                        ex.Message.IndexOf("protected sheet") < 0 &&
                        ex.Message.IndexOf("zamknutém listu") < 0) {
                        Assert.Fail();
                    }
                } finally {
                    exApp.DisplayAlerts = false;
                    axWb.Close();
                    exApp.Quit();
                }

                exApp = new Microsoft.Office.Interop.Excel.Application();
                exApp.Visible = true;
                axWb = exApp.Workbooks.Open(unlockFileName);
                axWs = axWb.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
                axRange = axWs.Range[strRange];
                try {
                    axRange.Value = "new value ssdggsd";
                } catch {
                    Assert.Fail();
                } finally {
                    exApp.DisplayAlerts = false;
                    axWb.Close();
                    exApp.Quit();
                }

                Assert.IsTrue(1 == 1);

            } catch (Exception ex) {
                throw ex;
            } finally {
                try {
                    File.Delete(excelFileName);
                    File.Delete(unlockFileName);
                } catch { }
            }
        }

        [TestMethod()]
        public void LockXlsmSheetPwd_Save_UnlockwoPwd() {
            string excelFileName = null;
            string unlockFileName = null;
            try {
                // Arrange
                Microsoft.Office.Interop.Excel.Application exApp = new Microsoft.Office.Interop.Excel.Application();
                exApp.Visible = true;
                var axWb = exApp.Workbooks.Add();
                var axWs = axWb.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;

                //set values
                string strRange = "B1:B2";
                var axRange = axWs.Range[strRange];
                axRange.Value = "ewa >wq tqw < t3532 d";

                axWs.Protect("Password");
                excelFileName = GetExcelXlsmFileName();
                axWb.SaveAs(excelFileName);
                axWb.Close();
                exApp.Quit();

                // Act
                unlockFileName = Excel.Unlock(excelFileName);

                // Assert
                exApp = new Microsoft.Office.Interop.Excel.Application();
                exApp.Visible = true;
                axWb = exApp.Workbooks.Open(excelFileName);
                axWs = axWb.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
                axRange = axWs.Range[strRange];
                try {
                    axRange.Value = "new value ssdggsd";
                } catch (Exception ex) {
                    if (ex.Message.IndexOf("The cell or chart that you are trying to change is protected and therefore read-only") < 0 &&
                        ex.Message.IndexOf("Pokoušíte ze změnit zamknutou buňku nebo zamknutý graf, které jsou proto jen pro čtení") < 0 &&
                        ex.Message.IndexOf("protected sheet") < 0 &&
                        ex.Message.IndexOf("zamknutém listu") < 0) {
                        Assert.Fail();
                    }
                } finally {
                    exApp.DisplayAlerts = false;
                    axWb.Close();
                    exApp.Quit();
                }

                exApp = new Microsoft.Office.Interop.Excel.Application();
                exApp.Visible = true;
                axWb = exApp.Workbooks.Open(unlockFileName);
                axWs = axWb.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
                axRange = axWs.Range[strRange];
                try {
                    axRange.Value = "new value ssdggsd";
                } catch {
                    Assert.Fail();
                } finally {
                    exApp.DisplayAlerts = false;
                    axWb.Close();
                    exApp.Quit();
                }

                Assert.IsTrue(1 == 1);

            } catch (Exception ex) {
                throw ex;
            } finally {
                try {
                    File.Delete(excelFileName);
                    File.Delete(unlockFileName);
                } catch { }
            }
        }

        [TestMethod()]
        public void RemoveVbaProtection() {
            // Excel.Unlock(@"C:\Temp\XlsTest\vbapassword.xlsm");
            Excel.Unlock(@"d:\eLogistics\FIN Params\Romania\RO_202105.xlsx");
        }

        [TestMethod()]
        public void GenerateExcelWorkbook_50StringColumns() {
            //Arrange
            System.Data.DataTable t = new System.Data.DataTable();
            for (int i = 0; i < 50; i++) {
                System.Data.DataColumn col = new System.Data.DataColumn("col" + i, typeof(string));
                t.Columns.Add(col);
            }

            for (int i = 0; i < 550; i++) {
                var newRow = t.NewRow();
                for (int j = 0; j < 50; j++) {
                    newRow["col" + j] = (i * j).ToString();
                }
                t.Rows.Add(newRow);
            }

            //Act
            string excelFileName = GetExcelFileName();
            using (var memStream = new Excel().GenerateExcelWorkbook(t)) {
                FileStream file = new FileStream(excelFileName, FileMode.Create, FileAccess.Write);
                memStream.WriteTo(file);
                file.Close();
                memStream.Close();
            }

            var exApp = new Microsoft.Office.Interop.Excel.Application();
            exApp.Visible = true;
            var axWb = exApp.Workbooks.Open(excelFileName);
            var axWs = axWb.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;

            string cellValue = null;
            
            try {
                cellValue = (axWs.Cells[18, 2] as Microsoft.Office.Interop.Excel.Range).Value;
                //string test = axWs.Cells[18, 2].Value.ToString();
            } catch {
                Assert.Fail();
            } finally {
                exApp.DisplayAlerts = false;
                axWb.Close();
                exApp.Quit();
            }

            //Assert
            Assert.IsTrue(cellValue == "16");
        }

        [TestMethod()]
        public void GenerateExcelWorkbook_50IntColumns() {
            //Arrange
            System.Data.DataTable t = new System.Data.DataTable();
            for (int i = 0; i < 50; i++) {
                System.Data.DataColumn col = new System.Data.DataColumn("col" + i, typeof(int));
                t.Columns.Add(col);
            }

            for (int i = 0; i < 550; i++) {
                var newRow = t.NewRow();
                for (int j = 0; j < 50; j++) {
                    newRow["col" + j] = (i * j);
                }
                t.Rows.Add(newRow);
            }

            //Act
            string excelFileName = GetExcelFileName();
            using (var memStream = new Excel().GenerateExcelWorkbook(t)) {
                FileStream file = new FileStream(excelFileName, FileMode.Create, FileAccess.Write);
                memStream.WriteTo(file);
                file.Close();
                memStream.Close();
            }

            var exApp = new Microsoft.Office.Interop.Excel.Application();
            exApp.Visible = true;
            var axWb = exApp.Workbooks.Open(excelFileName);
            var axWs = axWb.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;

            object cellValue = null;

            try {
                cellValue = (axWs.Cells[18, 2] as Microsoft.Office.Interop.Excel.Range).Value;
                //string test = axWs.Cells[18, 2].Value.ToString();
            } catch {
                Assert.Fail();
            } finally {
                exApp.DisplayAlerts = false;
                axWb.Close();
                exApp.Quit();
            }

            //Assert
            Assert.IsTrue(Convert.ToInt32(cellValue) == 16);
        }
        #endregion

        #region Methods
        private string GetExcelFileName() {
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

        private string GetExcelXlsmFileName() {
            string strPureName = "TmpExcelTest";
            string strFilePath = Path.Combine(GetTmpFolder(), strPureName + ".xlsx");
            int iIndex = 1;
            while (File.Exists(strFilePath)) {
                strFilePath = Path.Combine(GetTmpFolder(), strPureName + "_" + iIndex + ".xlsm");
                iIndex++;
            }


            return strFilePath;
        }

        private string GetTmpFolder() {
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

        private void DeleteExcelTestFolder() {

        }
        #endregion

        
    }
}