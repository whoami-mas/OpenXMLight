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
        private OpenXmlSpreadsheet.SheetData? _sheetData;
        private OpenXmlSpreadsheet.Columns? _columns;

        public string? Name
        {
            get => SheetXml.Name;
            set => SheetXml.Name = value;
        }
        public Cells Cells { get; private set; }
        public Rows Rows { get; private set; }



        internal OpenXmlSpreadsheet.Sheet SheetXml { get; private set; }
        internal OpenXmlPackaging.WorksheetPart WorksheetPart { get; set; }
        internal OpenXmlSpreadsheet.SheetData SheetDataXml
        {
            get
            {
                if(_sheetData == null)
                {
                    _sheetData = WorksheetPart.Worksheet.GetFirstChild<OpenXmlSpreadsheet.SheetData>();
                }

                return _sheetData;
            }
        }
        public OpenXmlSpreadsheet.Columns? Columns
        {
            get
            {
                if(_columns == null)
                {
                    _columns = WorksheetPart.Worksheet.GetFirstChild<OpenXmlSpreadsheet.Columns>() ??
                        WorksheetPart.Worksheet.InsertAt<OpenXmlSpreadsheet.Columns>(new OpenXmlSpreadsheet.Columns(), 1);
                }

                return _columns;
            }
        }

        internal Sheet(OpenXmlPackaging.WorksheetPart worksheetPart, string? name = null)
        {
            Create(worksheetPart: worksheetPart);

            this.Name = name;
        }

        internal Sheet(OpenXmlSpreadsheet.Sheet sheetXml,
            OpenXmlPackaging.WorksheetPart worksheetPart = default) => this.Create(sheetXml, worksheetPart);



        internal void Create(OpenXmlSpreadsheet.Sheet sheetXml = default,
                             OpenXmlPackaging.WorksheetPart worksheetPart = default)
        {
            this.SheetXml = sheetXml ?? new();
            this.WorksheetPart = worksheetPart;

            Cells = new Cells(this);
            Rows = new Rows(this);
        }
    }
}
