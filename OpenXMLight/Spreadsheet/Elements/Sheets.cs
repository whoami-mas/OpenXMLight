using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXMLight.Spreadsheet.ExcelContext;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using OpenXml = DocumentFormat.OpenXml;

namespace OpenXMLight.Spreadsheet.Elements
{
    public class Sheets : IEnumerable<Sheet>
    {
        public int Count => sheets.Count;
        public bool IsReadOnly => false;



        private List<Sheet> sheets;
        private readonly Context _context;


        internal OpenXmlPackaging.SpreadsheetDocument Excel { get; set; }
        internal OpenXmlSpreadsheet.Sheets SheetsXml => Excel.WorkbookPart.Workbook.Sheets;
        


        public Sheet this[int index]
        {
            get => sheets[index];
            set => sheets[index] = value;
        }


        internal Sheets(OpenXmlPackaging.SpreadsheetDocument excel, Context _context)
        {
            this.Excel = excel;
            this._context = _context;

            this.sheets = SheetsXml.Select(
                    s => new Sheet(
                                   (OpenXmlSpreadsheet.Sheet)s,
                                   _context,
                                   (OpenXmlPackaging.WorksheetPart)Excel.WorkbookPart.GetPartById(((OpenXmlSpreadsheet.Sheet)s).Id)
                    )
                ).ToList();
        }


        #region functions
        public void Add(string nameSheet)
        {
            OpenXmlPackaging.WorksheetPart worksheetPart = Excel.WorkbookPart.AddNewPart<OpenXmlPackaging.WorksheetPart>();
            worksheetPart.Worksheet = new OpenXmlSpreadsheet.Worksheet(
                new OpenXmlSpreadsheet.SheetDimension() { Reference = "A1" },
                //new OpenXmlSpreadsheet.Columns(),
                new OpenXmlSpreadsheet.SheetData());

            Sheet item = new Sheet(worksheetPart, _context, nameSheet);

            OpenXml.UInt32Value maxIdSheet = sheets.Select(s => s.SheetXml.SheetId).Cast<OpenXml.UInt32Value>().DefaultIfEmpty((OpenXml.UInt32Value)0).Max();
            item.SheetXml.SheetId = maxIdSheet + 1;
            item.SheetXml.Id = Excel.WorkbookPart.GetIdOfPart(item.WorksheetPart);
            item.SheetXml.Name = item.Name ?? $"Лист{maxIdSheet + 1}";
            SheetsXml.AppendChild(item.SheetXml);

            //item.Cells = new Cells(item.WorksheetPart);

            sheets.Add(item);
        }

        public void Clear()
        {
            SheetsXml.RemoveAllChildren<OpenXmlSpreadsheet.Sheet>();
            Excel.WorkbookPart.DeleteParts<OpenXmlPackaging.WorksheetPart>(Excel.WorkbookPart.WorksheetParts);
            
            this.Add("Лист1");
        }

        public bool Contains(Sheet item) => sheets.Any(a => a.SheetXml.Id == item.SheetXml.Id);

        public void CopyTo(Sheet[] array, int arrayIndex) => throw new NotImplementedException("CopyTo is not supported");

        public bool Remove(Sheet item)
        {
            OpenXmlSpreadsheet.Sheet removeSheet = SheetsXml.OfType<OpenXmlSpreadsheet.Sheet>().FirstOrDefault(f => f.Id == item.SheetXml.Id);

            Excel.WorkbookPart.DeletePart(Excel.WorkbookPart.GetPartById(removeSheet.Id));
            removeSheet.Remove();
            sheets.Remove(item);

            return true;
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        public IEnumerator<Sheet> GetEnumerator() => sheets.GetEnumerator();
        #endregion
    }
}
