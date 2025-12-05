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
    public class RowsRangeBase : RangeBase
    {
        internal int _rowFrom;
        internal int _rowTo;

        internal override OpenXmlPackaging.WorksheetPart WorksheetPart { get; init; }
        internal override OpenXmlPackaging.WorkbookPart WorkbookPart { get; init; }
        internal override OpenXmlSpreadsheet.SheetData SheetData => WorksheetPart.Worksheet.Elements<OpenXmlSpreadsheet.SheetData>().First();

        public int Count => RowsXml.Count;
        public int CountCell => RowsXml.Sum(s=> s.Elements<OpenXmlSpreadsheet.Cell>().Count());
        
        internal List<OpenXmlSpreadsheet.Row> RowsXml { get; init; }

        internal RowsRangeBase(OpenXmlPackaging.WorksheetPart worksheetPart, OpenXmlPackaging.WorkbookPart workbookPart)
        {
            this.WorkbookPart = workbookPart;
            this.WorksheetPart = worksheetPart;

            RowsXml = new();
        }

        internal void GetData()
        {
            RowsXml.Clear();

            for(int i = _rowFrom; i <= _rowTo; i++)
            {
                RowsXml.Add(
                    SheetData.Elements<OpenXmlSpreadsheet.Row>().FirstOrDefault(f => f.RowIndex == i) 
                        ?? SheetData.AppendChild(new OpenXmlSpreadsheet.Row() { RowIndex = Convert.ToUInt32(i)})
                    );
            }
        }
    }
}
