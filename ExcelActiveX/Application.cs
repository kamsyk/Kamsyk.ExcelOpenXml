
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kamsyk.ExcelOpenXml.ExcelActiveX {
    public class Application : BaseExcelObject {
        #region Properties
        private Workbooks m_Workbooks = new Workbooks();

        
        public Workbooks Workbooks {
            get { return m_Workbooks; }
        }
        #endregion

        #region Methods
        public void Quit() {
            
            if (CurrSpreadsheetDocument != null) {
                CurrSpreadsheetDocument.Close();
                CurrSpreadsheetDocument = null;
            }
        }
        #endregion
    }
}
