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

namespace OpenXMLight.Spreadsheet.Elements
{
    public class CellsRangeBase : RangeBase
    {
        private object? _value = null;


        internal int _row;
        internal int _col;
        internal int _rowTo;
        internal int _colTo;
        internal string? _addressCell;


        internal override Context Context => Context.Instance;
        internal override Sheet Sheet { get; }

        //internal override OpenXmlPackaging.WorksheetPart WorksheetPart { get; init; }
        internal override OpenXmlSpreadsheet.SheetData SheetData { get; }
        internal OpenXmlSpreadsheet.MergeCells? MergeCells { get; }
        internal OpenXmlSpreadsheet.Cell? CellXml { get; private set; }



        internal CellsRangeBase(Sheet sheet)
        {
            Sheet = sheet;
            SheetData = sheet.WorksheetPart.Worksheet.Elements<OpenXmlSpreadsheet.SheetData>().First();
            MergeCells = sheet.WorksheetPart.Worksheet.Elements<OpenXmlSpreadsheet.MergeCells>().FirstOrDefault();
        }


        public object? Value
        {
            get
            {
                return _value;
            }
            set 
            {
                ChangeCellValue(value);

                _value = value;
            }
        }



        internal void GetData()
        {
            OpenXmlSpreadsheet.Row rowFind = SheetData.Elements<OpenXmlSpreadsheet.Row>().FirstOrDefault(f => f.RowIndex == _row)
                ?? SheetData.AppendChild(new OpenXmlSpreadsheet.Row() { RowIndex = Convert.ToUInt32(_row) });

            CellXml = rowFind.Elements<OpenXmlSpreadsheet.Cell>()
                                    .FirstOrDefault(f => string.Equals(f.CellReference, _addressCell)) 
                                    ?? rowFind.AppendChild(new OpenXmlSpreadsheet.Cell() { CellReference = $"{HelperData.GetColumnByIndex(_col)}{_row}" });

            //GetCellValue();
        }

        internal void ChangeCellValue(object input)
        {
            if(CellXml == null)
            {
                OpenXmlSpreadsheet.Row rowFind = SheetData.Elements<OpenXmlSpreadsheet.Row>().FirstOrDefault(f => f.RowIndex == _row);
                if(rowFind == null)
                {
                    rowFind = new OpenXmlSpreadsheet.Row() { RowIndex = Convert.ToUInt32(_row) };
                    SheetData.AppendChild(rowFind);
                }

                CellXml = new OpenXmlSpreadsheet.Cell(
                    new OpenXmlSpreadsheet.CellValue()
                    ) { CellReference = _addressCell };
                
                rowFind.AppendChild(CellXml);
            }

            if (CellXml.CellValue == null)
                CellXml.CellValue = new OpenXmlSpreadsheet.CellValue();

            if (string.Equals("Int32", input.GetType().Name))
            {
                CellXml.CellValue.Text = input.ToString();
            }
            else if(string.Equals("String", input.GetType().Name))
            {
                CellXml.DataType = OpenXmlSpreadsheet.CellValues.SharedString;

                int index = Context.SharedStrings.AppendValue(input.ToString());

                CellXml.CellValue.Text = index.ToString();
            }
        }

        //internal void GetCellValue()
        //{
        //    if (CellXml == null)
        //        return;

        //    if(CellXml.CellValue == null)
        //    {
        //        _value = null;

        //        return;
        //    }

        //    if (CellXml.DataType != null && CellXml.DataType == OpenXmlSpreadsheet.CellValues.SharedString)
        //    {
        //        int index = int.Parse(CellXml.CellValue.Text);

        //        _value = Context.SharedStrings.GetValueOfIndex(index);
        //    }
        //    else if(CellXml.StyleIndex != null)
        //    {
        //        int indexStyle = Convert.ToInt32(CellXml.StyleIndex.Value);

        //        Context.Styles.GetFormatteCell(indexStyle);
        //    }
        //    else
        //        _value = CellXml.CellValue?.InnerText;
        //}


        public void Remove()
        {
            if (CellXml == null) 
                return;

            OpenXmlSpreadsheet.Row rowFind = SheetData.Elements<OpenXmlSpreadsheet.Row>().FirstOrDefault(f => f.RowIndex == _row);
            if (rowFind == null)
                return;

            if(CellXml.DataType != null && CellXml.DataType == OpenXmlSpreadsheet.CellValues.SharedString)
            {
                int index = int.Parse(CellXml.CellValue.Text);
                Context.SharedStrings.RemoveElementOfIndex(index);
            }

            CellXml.Remove();
            CellXml = null;
        }

        #region Merge cells
        internal void Merge()
        {
            if (MergeCells == null)
                Sheet.WorksheetPart.Worksheet.AppendChild(new OpenXmlSpreadsheet.MergeCells());
        }
        public void Merge(int rowTo, int colTo)
        {
            this.Merge();

            string addressMergeCell = $"{HelperData.GetColumnByIndex(colTo)}{rowTo}";
            string address = $"{_addressCell}:{addressMergeCell}";

            ValidationExcel.ValidationMerge(MergeCells, _row, _col, rowTo, colTo, address);
            
            for(int i = _row; i <= rowTo; i++)
            {
                OpenXmlSpreadsheet.Row rowFind = SheetData.Elements<OpenXmlSpreadsheet.Row>().FirstOrDefault(f => f.RowIndex == Convert.ToUInt32(i)) 
                    ?? SheetData.AppendChild(new OpenXmlSpreadsheet.Row() { RowIndex = Convert.ToUInt32(i)});

                for (int j = _col + 1; j <= colTo; j++)
                {
                    string appendAddress = $"{HelperData.GetColumnByIndex(j)}{i}";

                    OpenXmlSpreadsheet.Cell cell = rowFind.Elements<OpenXmlSpreadsheet.Cell>().FirstOrDefault(f => string.Equals(f.CellReference, addressMergeCell)) 
                        ?? rowFind.AppendChild(new OpenXmlSpreadsheet.Cell() { CellReference = appendAddress });
                }
            }

            MergeCells.AppendChild(
                new OpenXmlSpreadsheet.MergeCell() { Reference = address }
            );
        }
        public void Merge(string addressCellTo)
        {
            int rowTo = HelperData.GetRowIndex(addressCellTo);
            int colTo = HelperData.GetColumnIndex(addressCellTo);

            this.Merge(rowTo, colTo);
        }
        #endregion
    }
}
