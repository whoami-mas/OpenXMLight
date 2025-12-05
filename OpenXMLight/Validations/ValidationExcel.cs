using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLight.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLight.Validations
{
    internal static class ValidationExcel
    {
        internal static void ValidationIndexRow(int row)
        {
            if (row < 1 || row > 1048576)
                throw new ArgumentOutOfRangeException("Индекс строки неверен");
        }
        internal static void ValidationIndexColumn(int col)
        {
            if (col < 1 || col > 16384)
                throw new ArgumentOutOfRangeException("Индекс колонки неверен");
        }

        internal static void ValidationIndex(int row, int col)
        {
            ValidationIndexRow(row);
            ValidationIndexColumn(col);
        }



        internal static void ValidationAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
                throw new ArgumentNullException("Адрес не может быть пустым");

            Regex regex = new Regex(@"^[A-Z]+[0-9]+$", RegexOptions.IgnoreCase);
            if (!regex.IsMatch(address))
                throw new ArgumentException("Данные адрес не является валидным");

            int indexColumn = HalperData.GetColumnIndex(address);
            int indexRow = HalperData.GetRowIndex(address);
            ValidationIndex(indexRow, indexColumn);
        }



        internal static void ValidationMerge(int rowFrom, int colFrom, int rowTo, int colTo)
        {
            if (rowTo < rowFrom || colTo < colFrom)
                throw new ArgumentException("Индекс не является валидным");
        }
        internal static void ValidationMerge(OpenXmlSpreadsheet.MergeCells mergeCells, int rowFrom, int colFrom, int rowTo, int colTo, string findAddress)
        {
            ValidationMerge(rowFrom, colFrom, rowTo, colTo);
            ValidationMerge(mergeCells, findAddress);

            foreach (OpenXmlSpreadsheet.MergeCell item in mergeCells.ChildElements.Cast<OpenXmlSpreadsheet.MergeCell>())
            {
                string[] address = item.Reference.ToString().Split(":");
                
                int indexMinRow = HalperData.GetRowIndex(address[0]);
                int indexMaxRow = HalperData.GetRowIndex(address[1]);

                int indexMinCol = HalperData.GetRowIndex(address[0]);
                int indexMaxCol = HalperData.GetRowIndex(address[1]);

                bool isRangeCellFrom = rowFrom >= indexMinRow && rowFrom <= indexMaxRow &&
                    colFrom >= indexMinCol && colFrom <= indexMaxCol;

                if (isRangeCellFrom)
                    throw new ArgumentException($"Ячейка {HalperData.GetColumnByIndex(colFrom)}{rowFrom} уже объеденена");
            }
        }
        internal static void ValidationMerge(OpenXmlSpreadsheet.MergeCells mergeCells,  string findAddress)
        {
            if (
                mergeCells.Any(a =>
                {
                    var addressMerge = (OpenXmlSpreadsheet.MergeCell)a;
                    return string.Equals(addressMerge, findAddress);
                })
              )
                throw new ArgumentException("Данный адрес объединения уже существует");
        }
        
    }
}
