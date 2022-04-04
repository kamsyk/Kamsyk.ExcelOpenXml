using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kamsyk.ExcelOpenXml.ExcelActiveX {
    public class Worksheets : BaseNotInstantiable {
        #region Properties
        private WorkbookPart m_WorkbookPart = null;
        private List<Workbook> m_Items = null;
        public int Count {
            get {
                if (m_Items == null) {
                    return 0;
                }

                return m_Items.Count;
            }
        }
        #endregion

        #region Methods
        
        #endregion

       
    }
}
