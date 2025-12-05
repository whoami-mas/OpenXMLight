using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLight.Spreadsheet.Elements
{
    public abstract class RangeBase
    {
        internal abstract OpenXmlSpreadsheet.SheetData SheetData { get; }
        internal abstract OpenXmlPackaging.WorksheetPart WorksheetPart { get; init; }
        internal abstract OpenXmlPackaging.WorkbookPart WorkbookPart { get; init; }

    }
}
