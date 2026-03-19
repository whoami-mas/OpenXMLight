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


        internal override Context Context => Sheet._context;
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
            // Инициализация MergeCells
            if (MergeCells == null)
                MergeCells = Sheet.WorksheetPart.Worksheet.AppendChild<OpenXmlSpreadsheet.MergeCells>(
                    new OpenXmlSpreadsheet.MergeCells());

            string addressMergeCell = $"{HelperData.GetColumnByIndex(_colTo)}{_rowTo}";

            // Валидация
            ValidationExcel.ValidationMerge(MergeCells, _row, _col, _rowTo, _colTo, _addressCell);

            OpenXmlSpreadsheet.Cell? firstCell = null;

            // Проходим по всем ячейкам диапазона
            for (int i = _row; i <= _rowTo; i++)
            {
                // Получаем или создаем строку
                OpenXmlSpreadsheet.Row rowFind = SheetData.Elements<OpenXmlSpreadsheet.Row>()
                    .FirstOrDefault(f => f.RowIndex == (uint)i)
                    ?? SheetData.AppendChild(new OpenXmlSpreadsheet.Row() { RowIndex = (uint)i });
                
                for (int j = _col; j <= _colTo; j++)
                {
                    string cellAddress = $"{HelperData.GetColumnByIndex(j)}{i}";

                    // Получаем или создаем ячейку
                    OpenXmlSpreadsheet.Cell cell = rowFind.Elements<OpenXmlSpreadsheet.Cell>()
                        .FirstOrDefault(f => string.Equals(f.CellReference, cellAddress))
                        ?? rowFind.AppendChild(new OpenXmlSpreadsheet.Cell() { CellReference = cellAddress, StyleIndex = 0 });
                    
                    
                     // First cell ?  - Save
                    if (i == _row && j == _col)
                    {
                        firstCell = cell;
                        continue;
                    }

                    // Если первая ячейка пуста, а текущая имеет значение - копируем
                    if (firstCell?.CellValue == null && cell.CellValue != null &&
                        !string.IsNullOrEmpty(cell.CellValue.Text))
                    {
                        firstCell.CellValue = (OpenXmlSpreadsheet.CellValue)cell.CellValue.CloneNode(true);
                        firstCell.StyleIndex = cell.StyleIndex;
                        firstCell.DataType = cell.DataType;
                    }

                    // Применяем стиль первой ячейки
                    if (firstCell.StyleIndex != null && firstCell.StyleIndex.HasValue)
                    {
                        cell.StyleIndex = firstCell.StyleIndex.Value;
                    }

                    // Очищаем значение во всех ячейках кроме первой
                    if (cell != firstCell)
                    {
                        cell.CellValue?.RemoveAllChildren();
                        cell.CellValue = null;
                    }
                }
            }

            Sorted();

            MergeCells.AppendChild(
               new OpenXmlSpreadsheet.MergeCell() { Reference = _addressCell }
           );
        }
        #endregion

        private void Sorted()
        {
            foreach (var row in SheetData.Elements<OpenXmlSpreadsheet.Row>())
            {
                SortRow(row);
            }
        }
        private void SortRow(OpenXmlSpreadsheet.Row row)
        {
            var cells = row.Elements<OpenXmlSpreadsheet.Cell>().ToList();

            // Проверяем, нужно ли сортировать
            bool needSort = false;
            int prevIndex = 0;

            foreach (var cell in cells)
            {
                int currentIndex = HelperData.GetColumnIndex(cell.CellReference);
                if (currentIndex < prevIndex)
                {
                    needSort = true;
                    break;
                }
                prevIndex = currentIndex;
            }

            if (needSort)
            {
                // Сортируем ячейки по индексу колонки
                var sortedCells = cells.OrderBy(c => HelperData.GetColumnIndex(c.CellReference)).ToList();

                // Удаляем все ячейки
                foreach (var cell in cells)
                    cell.Remove();

                // Добавляем в правильном порядке
                foreach (var cell in sortedCells)
                    row.AppendChild(cell);
            }
        }

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
