using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXMLight.Spreadsheet.Parts;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLight.Spreadsheet.ExcelContext
{
    internal class Context: IContext
    {
        public Styles Styles { get; init; }
        public SharedStrings SharedStrings { get; init; }


        internal Context(OpenXmlPackaging.WorkbookPart workbookPart)
        {
            if(workbookPart != null)
            {
                Styles = new(workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<OpenXmlPackaging.WorkbookStylesPart>());
                SharedStrings = new(workbookPart.SharedStringTablePart ?? workbookPart.AddNewPart<OpenXmlPackaging.SharedStringTablePart>());
            }
        }
    }
}
