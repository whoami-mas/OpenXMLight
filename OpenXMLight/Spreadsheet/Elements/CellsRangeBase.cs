using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXMLight.Spreadsheet.ExcelContext;
using OpenXMLight.Tools;
using OpenXMLight.Validations;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

using OpenXMLight.Spreadsheet.Formatting;

namespace OpenXMLight.Spreadsheet.Elements
{
    public class CellsRangeBase : RangeBase
    {
        internal int _row;
        internal int _col;
        internal int _rowTo;
        internal int _colTo;
        internal string? _addressCell;


        internal override Context Context => Context.Instance;
        internal override Sheet Sheet { get; }
        internal override OpenXmlSpreadsheet.SheetData SheetData { get; }
        internal OpenXmlSpreadsheet.MergeCells? MergeCells { get; private set; }

        internal CellsRangeBase(Sheet sheet)
        {
            Sheet = sheet;
            SheetData = sheet.WorksheetPart.Worksheet.Elements<OpenXmlSpreadsheet.SheetData>().First();
            MergeCells = sheet.WorksheetPart.Worksheet.Elements<OpenXmlSpreadsheet.MergeCells>().FirstOrDefault();
        }

        #region Merge cells
        public void Merge()
        {
            //this.Merge();
            if (MergeCells == null)
                MergeCells = Sheet.WorksheetPart.Worksheet.AppendChild<OpenXmlSpreadsheet.MergeCells>(new OpenXmlSpreadsheet.MergeCells());

            string addressMergeCell = $"{HelperData.GetColumnByIndex(_colTo)}{_rowTo}";
            //string address = $"{_addressCell}:{addressMergeCell}";

            ValidationExcel.ValidationMerge(MergeCells, _row, _col, _rowTo, _colTo, _addressCell);
            
            for(int i = _row; i <= _rowTo; i++)
            {
                OpenXmlSpreadsheet.Row rowFind = SheetData.Elements<OpenXmlSpreadsheet.Row>().FirstOrDefault(f => f.RowIndex == Convert.ToUInt32(i)) 
                    ?? SheetData.AppendChild(new OpenXmlSpreadsheet.Row() { RowIndex = Convert.ToUInt32(i)});

                for (int j = _col + 1; j <= _colTo; j++)
                {
                    string appendAddress = $"{HelperData.GetColumnByIndex(j)}{i}";

                    OpenXmlSpreadsheet.Cell cell = rowFind.Elements<OpenXmlSpreadsheet.Cell>().FirstOrDefault(f => string.Equals(f.CellReference, addressMergeCell)) 
                        ?? rowFind.AppendChild(new OpenXmlSpreadsheet.Cell() { CellReference = appendAddress });
                }
            }

            MergeCells.AppendChild(
                new OpenXmlSpreadsheet.MergeCell() { Reference = _addressCell }
            );
        }
        #endregion

        #region Styles
        public void SetFont(Action<Font> conf)
        {
            for(int i = _row; i <= _rowTo; i++)
            {
                for(int j = _col; j <= _colTo; j++)
                {
                    var cell = new Cell(Sheet, i, j);
                    
                    conf.Invoke(cell.Style.Font);

                    cell.Style.Font.CommitChange();
                }
            }
        }

        public void SetBorder(Action<Border> conf)
        {
            for (int i = _row; i <= _rowTo; i++)
            {
                for (int j = _col; j <= _colTo; j++)
                {
                    var cell = new Cell(Sheet, i, j);

                    conf.Invoke(cell.Style.Borders);

                    cell.Style.Borders.CommitChange();
                }
            }
        }

        public void SetWrapText(bool wrapText)
        {
            for (int i = _row; i <= _rowTo; i++)
            {
                for (int j = _col; j <= _colTo; j++)
                {
                    var cell = new Cell(Sheet, i, j);

                    cell.Style.IsWrap = wrapText;
                }
            }
        }
        public void SetHorizontalAlignment(HorizontalAlignments alignment)
        {
            for (int i = _row; i <= _rowTo; i++)
            {
                for (int j = _col; j <= _colTo; j++)
                {
                    var cell = new Cell(Sheet, i, j);

                    cell.Style.Horizontal = alignment;
                }
            }
        }
        public void SetVerticalAlignment(VerticalAlignments alignment)
        {
            for (int i = _row; i <= _rowTo; i++)
            {
                for (int j = _col; j <= _colTo; j++)
                {
                    var cell = new Cell(Sheet, i, j);

                    cell.Style.Vertical = alignment;
                }
            }
        }
        #endregion
    }
}
