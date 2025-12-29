using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXMLight.Tools;
using OpenXMLight.Validations;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using OpenXml = DocumentFormat.OpenXml;

namespace OpenXMLight.Spreadsheet.Elements
{
    public class Cells : CellsRangeBase
    {
        public Cells this[int row, int col]
        {
            get
            {
                ValidationExcel.ValidationIndex(row, col);

                _row = row;
                _col = col;
                _addressCell = $"{HelperData.GetColumnByIndex(_col)}{_row}";

                GetData();

                return this;
            }
        }
        public Cells this[string address]
        {
            get
            {
                ValidationExcel.ValidationAddress(address);
                _row = HelperData.GetRowIndex(address);
                _col = HelperData.GetColumnIndex(address);
                _addressCell = address;

                GetData();

                return this;
            }
        }

        internal Cells(OpenXmlPackaging.WorksheetPart worksheetPart, OpenXmlPackaging.WorkbookPart workbookPart)
            : base(worksheetPart, workbookPart)
        {

        }

    }
}
