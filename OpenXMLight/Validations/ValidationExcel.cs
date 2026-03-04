using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLight.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
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
        internal static void ValidationIndex(int row, int col, int rowTo, int colTo)
        {
            ValidationIndexRow(row);
            ValidationIndexColumn(col);

            if (row > rowTo || col > colTo)
                throw new ArgumentOutOfRangeException("Неверный диапазон ячеек");
        }


        internal static void ValidationAddress(string address)
        {
            Regex regex = new Regex(@"^[A-Z]+[0-9]+$", RegexOptions.IgnoreCase);
            if (!regex.IsMatch(address))
                throw new ArgumentException("Данные адрес не является валидным");

            int indexColumn = HelperData.GetColumnIndex(address);
            int indexRow = HelperData.GetRowIndex(address);
            ValidationIndex(indexRow, indexColumn);
        }
        internal static void ValidationFullAddress(string fullAddress, ref int rowFrom, ref int colFrom, ref int rowTo, ref int colTo)
        {
            if (string.IsNullOrWhiteSpace(fullAddress))
                throw new ArgumentNullException("Адрес не может быть пустым");

            string[] addresses = fullAddress.Split(':');

            foreach(var address in addresses)
                ValidationAddress(address);

            colFrom = HelperData.GetColumnIndex(addresses[0]);
            rowFrom = HelperData.GetRowIndex(addresses[0]);
            ValidationIndex(rowFrom, colFrom);

            if (addresses.Count() > 1)
            {
                colTo = HelperData.GetColumnIndex(addresses[1]);
                rowTo = HelperData.GetRowIndex(addresses[1]);
                ValidationIndex(rowTo, colTo);
            }
            else
            {
                colTo = colFrom;
                rowTo = rowFrom;
            }
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
                
                int indexMinRow = HelperData.GetRowIndex(address[0]);
                int indexMaxRow = HelperData.GetRowIndex(address[1]);

                int indexMinCol = HelperData.GetRowIndex(address[0]);
                int indexMaxCol = HelperData.GetRowIndex(address[1]);

                bool isRangeCellFrom = rowFrom >= indexMinRow && rowFrom <= indexMaxRow &&
                    colFrom >= indexMinCol && colFrom <= indexMaxCol;

                if (isRangeCellFrom)
                    throw new ArgumentException($"Ячейка {HelperData.GetColumnByIndex(colFrom)}{rowFrom} уже объеденена");
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
