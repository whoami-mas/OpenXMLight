using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.XPath;
using elements = OpenXMLight.Spreadsheet.Elements;

namespace OpenXMLight
{
    public class ExcelDocument : IDisposable
    {
        private SpreadsheetDocument? ExcelDoc { get; set; }
        public elements.Sheets Sheets { get; private set; }

        #region Dispose
        public void Dispose()
        {
            ExcelDoc?.Dispose();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion

        public void Save()
        {
            ExcelDoc?.WorkbookPart.Workbook.Save();
            ExcelDoc?.Dispose();
        }

        public ExcelDocument(string path, bool overwrite = false)
        {
            if (overwrite)
                File.Delete(path);

            ExcelDoc = File.Exists(path) ? SpreadsheetDocument.Open(path, true) 
                                         : SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);

            if (ExcelDoc.WorkbookPart == null)
            {
                ExcelDoc.AddWorkbookPart().Workbook = new Workbook();
                ExcelDoc.WorkbookPart.AddNewPart<SharedStringTablePart>().SharedStringTable = new SharedStringTable();
                ExcelDoc.WorkbookPart.Workbook.AppendChild(new Sheets());
            }

            Sheets = new elements.Sheets(ExcelDoc);

            if (Sheets.Count < 1)
                Sheets.Add("Лист1");
        }
    }
}
