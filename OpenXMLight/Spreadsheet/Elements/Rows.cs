using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using OpenXml = DocumentFormat.OpenXml;
using OpenXMLight.Validations;

namespace OpenXMLight.Spreadsheet.Elements
{
    public class Rows : RowsRangeBase
    {
        public Rows this[int row]
        {
            get
            {
                ValidationExcel.ValidationIndexRow(row);
                
                _rowFrom = row;
                _rowTo = row;

                GetData();

                return this;
            }
        }

        public Rows this[int rowFrom, int rowTo]
        {
            get
            {
                ValidationExcel.ValidationIndexRow(rowFrom);
                ValidationExcel.ValidationIndexRow(rowTo);

                _rowFrom = rowFrom;
                _rowTo = rowTo;

                GetData();

                return this;
            }
        }

        public int Count => SheetData.Elements<OpenXmlSpreadsheet.Row>().Count();

        internal Rows(OpenXmlPackaging.WorksheetPart worksheetPart, OpenXmlPackaging.WorkbookPart workbookPart) 
            : base(worksheetPart, workbookPart)
        {

        }
    }
}
