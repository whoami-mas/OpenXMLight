using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLight.Spreadsheet.ExcelContext;
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


        private string _tmp_path;


        public elements.Sheets Sheets { get; private set; }
        private Context Context { get; init;}
        public string FullPath => Path.GetFullPath(_tmp_path);

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
                ExcelDoc.AddWorkbookPart().Workbook = new Workbook(new Sheets());

            Context = new Context(ExcelDoc.WorkbookPart);

            Sheets = new elements.Sheets(ExcelDoc, Context);

            if (Sheets.Count < 1)
                Sheets.Add("Лист1");

            _tmp_path = path;
        }
    }
}
