using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kamsyk.ExcelOpenXml.ExcelActiveX {
    public class Cell : Range {
        #region Properties
        internal DocumentFormat.OpenXml.Spreadsheet.Cell m_Cell = null;
        #endregion

        #region Constructor
        internal Cell(Worksheet worksheet, int iRowStart, int iColStart, int iRowEnd, int iColEnd) : 
            base(worksheet, iRowStart, iColStart, iRowEnd, iColEnd) {
        }

        internal Cell(Worksheet worksheet, int iRow, int iCol) :
            base(worksheet, iRow, iCol, iRow, iCol) {
        }
        #endregion

        #region Methods
        internal void SetCell(DocumentFormat.OpenXml.Spreadsheet.Cell xmlCell) {
            m_Cell = xmlCell;
        }
        #endregion
    }
}
