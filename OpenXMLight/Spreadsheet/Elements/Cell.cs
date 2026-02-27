using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using OpenXml = DocumentFormat.OpenXml;

using OpenXMLight.Spreadsheet.Formatting;
using OpenXMLight.Tools;
using OpenXMLight.Spreadsheet.ExcelContext;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLight.Spreadsheet.Elements
{
    public class Cell
    {
        internal OpenXmlSpreadsheet.Cell? cellXml;
        internal Sheet _sheet;

        private object? _value = null;
        private TypeValue _typeValue;
        private int _row;
        private int _col;
        private StyleCell _style;

        private Context Context => Context.Instance;

        public object? Value
        {
            get
            {
                GetCellValue();

                return _value;
            }
            set
            {
                if(cellXml == null)
                    GetCreateCell();
                
                ChangeCellValue(value);
            }
        }
        public StyleCell Style
        {
            get => _style;
        }

        internal Cell(Sheet sheet, int _row, int _col)
        {
            this._sheet = sheet;
            this._row = _row;
            this._col = _col;

            GetCellXml();
        }


        private void GetCellXml()
        {
            OpenXmlSpreadsheet.Row? rowXml = _sheet.SheetDataXml.Elements<OpenXmlSpreadsheet.Row>().FirstOrDefault(f => f.RowIndex == Convert.ToUInt32(_row));

            if(rowXml != null)
                this.cellXml = rowXml.Elements<OpenXmlSpreadsheet.Cell>().FirstOrDefault(f => string.Equals(f.CellReference, $"{HelperData.GetColumnByIndex(_col)}{_row}"));
        
            _style = new StyleCell(cellXml);
        }
        private void GetCreateCell()
        {
            OpenXmlSpreadsheet.Row? rowXml = _sheet.SheetDataXml.Elements<OpenXmlSpreadsheet.Row>().FirstOrDefault(f => f.RowIndex == Convert.ToUInt32(_row))
                ?? _sheet.SheetDataXml.AppendChild(new OpenXmlSpreadsheet.Row() { RowIndex = Convert.ToUInt32(_row) });

            this.cellXml = rowXml.AppendChild(new OpenXmlSpreadsheet.Cell() { CellReference = $"{HelperData.GetColumnByIndex(_col)}{_row}" });
        }

        internal void GetCellValue()
        {
            if (cellXml == null || cellXml.CellValue == null)
                return;

            if (cellXml.DataType != null && cellXml.DataType == OpenXmlSpreadsheet.CellValues.SharedString)
            {
                int index = int.Parse(cellXml.CellValue.Text);

                _value = Context.SharedStrings.GetValueOfIndex(index);
            }
            else if (cellXml.StyleIndex != null)
            {
                int indexStyle = Convert.ToInt32(cellXml.StyleIndex.Value);

                TypeValue format = Context.Styles.GetFormatteCell(indexStyle);

                switch (format)
                {
                    case var f when f == TypeValue.Date:
                        Style.Type = TypeValue.Date;
                        _value = DateTime.FromOADate(Convert.ToInt32(cellXml.CellValue.Text));
                        break;
                    case var f when f == TypeValue.Number:
                        Style.Type = TypeValue.Number;
                        _value = Convert.ToInt32(cellXml.CellValue.Text);
                        break;
                    case var f when f == TypeValue.Percent:
                        Style.Type = TypeValue.Percent;
                        _value = cellXml.CellValue.Text;
                        break;
                    case var f when f == TypeValue.General:
                        Style.Type = TypeValue.General;
                        _value = cellXml.CellValue.Text;
                        break;
                    case var f when f == TypeValue.Other:
                        Style.Type = TypeValue.Other;
                        _value = cellXml.CellValue.Text;
                        break;
                }
            }
            else
                _value = cellXml.CellValue?.InnerText;
        }
        internal void ChangeCellValue(object? input)
        {
            if (cellXml.CellValue == null)
                cellXml.CellValue = new OpenXmlSpreadsheet.CellValue();

            if (string.Equals("Int32", input.GetType().Name))
            {
                cellXml.CellValue.Text = input.ToString();
            }
            else if(string.Equals("DateTime", input.GetType().Name))
            {
                if (!DateTime.TryParse(input.ToString(), out DateTime date))
                    throw new ArgumentException("дата не может быть пустой");

                cellXml.CellValue.Text = (date.ToOADate()).ToString();
                cellXml.StyleIndex = Convert.ToUInt32(Context.Styles.GetFormatteCellIndex(TypeValue.Date));
            }
            else if (string.Equals("String", input.GetType().Name))
            {
                cellXml.DataType = OpenXmlSpreadsheet.CellValues.SharedString;

                int index = Context.SharedStrings.AppendValue(input.ToString());

                cellXml.CellValue.Text = index.ToString();
            }
        }
    }
}
