using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using OpenXml = DocumentFormat.OpenXml;

namespace OpenXMLight.Spreadsheet.Elements
{
    public class Sheet
    {
        public string? Name
        {
            get => SheetXml.Name;
            set => SheetXml.Name = value;
        }
        public Cells Cells { get; private set; }

        internal OpenXmlSpreadsheet.Sheet SheetXml { get; private set; }
        internal OpenXmlPackaging.WorksheetPart WorksheetPart { get; set; }
        internal OpenXmlPackaging.WorkbookPart WorkbookPart { get; set; }
        
        internal Sheet(OpenXmlPackaging.WorkbookPart workbookPart, OpenXmlPackaging.WorksheetPart worksheetPart, string? name = null)
        {
            Create(workbookPart, worksheetPart: worksheetPart);

            this.Name = name;
        }

        
        internal Sheet(OpenXmlSpreadsheet.Sheet sheetXml,
            OpenXmlPackaging.WorkbookPart workbookPart,
            OpenXmlPackaging.WorksheetPart worksheetPart = default) => this.Create(workbookPart, sheetXml, worksheetPart);


        internal void Create(OpenXmlPackaging.WorkbookPart workbookPart,
                             OpenXmlSpreadsheet.Sheet sheetXml = default,
                             OpenXmlPackaging.WorksheetPart worksheetPart = default)
        {
            SheetXml = sheetXml ?? new();
            this.WorksheetPart = worksheetPart;
            this.WorkbookPart = workbookPart;

            Cells = new Cells(this.WorksheetPart, this.WorkbookPart);
        }
    }
}
