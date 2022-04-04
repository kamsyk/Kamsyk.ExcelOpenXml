using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kamsyk.ExcelOpenXml.ExcelActiveX {
    public class Workbook : BaseNotInstantiable {
        #region Propereties
        private WorkbookPart m_WorkbookPart = null;
        public List<Worksheet> Worksheets {
            get { return GetSheets(); }
        }
        #endregion

        #region Methods
        internal void SetWorkbookPart(WorkbookPart workbookPart) {
            m_WorkbookPart = workbookPart;
        }

        private List<Worksheet> GetSheets() {
            List<Worksheet> sheets = new List<Worksheet>();
            foreach (var sheet in m_WorkbookPart.Workbook.Sheets) {
                DocumentFormat.OpenXml.Spreadsheet.Sheet xmlSheet = (DocumentFormat.OpenXml.Spreadsheet.Sheet)sheet;
                Worksheet ws = new Worksheet();
                ws.SetSheet(xmlSheet);
                sheets.Add(ws);
            }

            return sheets;
        }
        #endregion
    }
}
