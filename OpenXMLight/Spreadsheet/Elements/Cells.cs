using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXMLight.Tools;
using OpenXMLight.Validations;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using OpenXml = DocumentFormat.OpenXml;

namespace OpenXMLight.Spreadsheet.Elements
{
    public class Cells : CellsRangeBase
    {
        public Cells this[int row, int col]
        {
            get
            {
                ValidationExcel.ValidationIndex(row, col);

                _row = row;
                _col = col;
                _addressCell = $"{HelperData.GetColumnByIndex(_col)}{_row}";

                GetData();

                return this;
            }
        }
        public Cells this[string address]
        {
            get
            {
                ValidationExcel.ValidationAddress(address);
                _row = HelperData.GetRowIndex(address);
                _col = HelperData.GetColumnIndex(address);
                _addressCell = address;

                GetData();

                return this;
            }
        }

        internal Cells(OpenXmlPackaging.WorksheetPart worksheetPart, OpenXmlPackaging.WorkbookPart workbookPart)
            : base(worksheetPart, workbookPart)
        {

        }


        #region AutoFitColumns

        public void AutoFitColumns()
        {
            try
            {
                var rows = SheetData.Elements<OpenXmlSpreadsheet.Row>();
                if (!rows.Any()) return;

                List<int> indexColumns = new();

                OpenXmlSpreadsheet.Columns columns = WorksheetPart.Worksheet.GetFirstChild<OpenXmlSpreadsheet.Columns>()
                   ?? WorksheetPart.Worksheet.InsertAt<OpenXmlSpreadsheet.Columns>(new OpenXmlSpreadsheet.Columns(), 1);
                columns.RemoveAllChildren<OpenXmlSpreadsheet.Column>();

                foreach (var row in rows)
                {
                    var cells = row.Elements<OpenXmlSpreadsheet.Cell>();

                    foreach (var cell in cells)
                    {
                        int indexRow = Convert.ToInt32(row.RowIndex.Value);
                        int indexCell = HelperData.GetColumnIndex(cell.CellReference);

                        var column = columns.Elements<OpenXmlSpreadsheet.Column>().FirstOrDefault(f => indexCell >= f.Min && indexCell <= f.Max)
                            ?? columns.AppendChild(new OpenXmlSpreadsheet.Column() 
                                {
                                    Min = Convert.ToUInt32(indexCell),
                                    Max = Convert.ToUInt32(indexCell),
                                    BestFit = true,
                                    CustomWidth = true
                                });

                        double width = HelperData.GetMaxLengthWidthCell(this[indexRow, indexCell].Value.ToString());

                        if(column.Width == null || column.Width == 0 || column.Width < width)
                            column.Width = width;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка определения автоматической ширины {ex.Message}");
            }
        }

        #endregion
    }
}
