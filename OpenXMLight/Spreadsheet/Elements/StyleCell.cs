using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using OpenXml = DocumentFormat.OpenXml;

using OpenXMLight.Spreadsheet.Formatting;
using OpenXMLight.Tools;
using OpenXMLight.Spreadsheet.ExcelContext;

namespace OpenXMLight.Spreadsheet.Elements
{
    public class StyleCell
    {
        internal OpenXmlSpreadsheet.Cell? cellXml;


        private TypeValue _typeValue = TypeValue.General;

        public TypeValue Type
        {
            get => _typeValue;
            set
            {

            }
        }


        internal StyleCell(OpenXmlSpreadsheet.Cell? cellXml)
        {
            this.cellXml = cellXml;
        }
    }
}
