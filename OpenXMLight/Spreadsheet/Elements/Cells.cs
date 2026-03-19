using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXMLight.Tools;
using OpenXMLight.Validations;
using OpenXMLight.Spreadsheet.Elements;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using OpenXml = DocumentFormat.OpenXml;

namespace OpenXMLight.Spreadsheet.Elements
{
    public class Cells : CellsRangeBase
    {
        public Cell this[int row, int col]
        {
            get
            {
                ValidationExcel.ValidationIndex(row, col);
                _row = row;
                _col = col;
                _rowTo = row;
                _colTo = col;

                return new Cell(Sheet, row, col);
            }
        }
        public Cells this[int rowFrom, int colFrom, int rowTo, int colTo]
        {
            get
            {
                ValidationExcel.ValidationIndex(rowFrom, colFrom, rowTo, colTo);
                _row = rowFrom;
                _col = colFrom;
                _rowTo = rowTo;
                _colTo = colTo;
                _addressCell = $"{HelperData.GetColumnByIndex(_col)}{_row}:{HelperData.GetColumnByIndex(_colTo)}{_rowTo}";

                return this;
            }
        }
        public Cells this[string address]
        {
            get
            {
                ValidationExcel.ValidationFullAddress(address, ref _row, ref _col, ref _rowTo, ref _colTo);
                _addressCell = address;

                return this;
            }
        }

        internal Cells(Sheet sheet)
            : base(sheet)
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


                foreach (var row in rows)
                {
                    var cells = row.Elements<OpenXmlSpreadsheet.Cell>();

                    foreach (var cell in cells)
                    {
                        int indexRow = Convert.ToInt32(row.RowIndex.Value);
                        int indexCell = HelperData.GetColumnIndex(cell.CellReference);

                        var column = Sheet.Columns.Elements<OpenXmlSpreadsheet.Column>().FirstOrDefault(f => indexCell >= f.Min && indexCell <= f.Max)
                            ?? Sheet.Columns.AppendChild(new OpenXmlSpreadsheet.Column() 
                                {
                                    Min = Convert.ToUInt32(indexCell),
                                    Max = Convert.ToUInt32(indexCell),
                                    BestFit = true,
                                    CustomWidth = true
                                });

                        object valueCell = this[indexRow, indexCell].Value;

                        if (valueCell == null)
                            continue;

                        double width = HelperData.GetMaxLengthWidthCell(valueCell.ToString());

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
