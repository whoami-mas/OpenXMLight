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

namespace OpenXMLight.Spreadsheet.Elements
{
    public class StyleCell
    {
        internal Context Context => _cell.Context;


        internal Cell _cell;

        internal OpenXmlSpreadsheet.Cell? cellXml => _cell.cellXml;
        internal uint? styleIndexCell => cellXml?.StyleIndex;
        internal OpenXmlSpreadsheet.CellFormat? cellFormat => Context.Styles.GetFormatteCell(cellXml?.StyleIndex);


        private TypeValue _typeValue = TypeValue.General;
        private HorizontalAlignments? _hAlignment;
        private VerticalAlignments? _vAlignment;
        private int _textRotation;
        private Font _font;
        private Border _borders;
        private bool _isWrap;

        public TypeValue Type
        {
            get => _typeValue;
            set
            {

            }
        }
        public HorizontalAlignments? Horizontal
        {
            get
            {
                if(_hAlignment == null)
                {
                    _hAlignment = HorizontalAlignments.Parse(cellFormat.Alignment?.Horizontal);
                }

                return _hAlignment;
            }
            set
            {
                cellXml.StyleIndex = Context.Styles.IsFirstFormatteCell(cellXml?.StyleIndex);

                cellFormat.Alignment ??= new OpenXmlSpreadsheet.Alignment();
                cellFormat.Alignment.Horizontal = value.Value.Value;
                cellFormat.ApplyAlignment = true;

                _hAlignment = value;
            }
        }
        public VerticalAlignments? Vertical
        {
            get
            {
                if(_vAlignment == null)
                {
                    _vAlignment = VerticalAlignments.Parse(cellFormat.Alignment?.Vertical);
                }

                return _vAlignment;
            }
            set
            {
                cellXml.StyleIndex = Context.Styles.IsFirstFormatteCell(cellXml?.StyleIndex);

                cellFormat.Alignment ??= new OpenXmlSpreadsheet.Alignment();
                cellFormat.Alignment.Vertical = value.Value.Value;
                cellFormat.ApplyAlignment = true;

                _vAlignment = value;
            }
        }
        public int TextRotation
        {
            get
            {
                _textRotation = (int)cellFormat.Alignment?.TextRotation.Value;

                return _textRotation;
            }
            set
            {
                cellXml.StyleIndex = Context.Styles.IsFirstFormatteCell(cellXml?.StyleIndex);

                _textRotation = value;

                cellFormat.Alignment ??= new OpenXmlSpreadsheet.Alignment();
                cellFormat.Alignment.TextRotation = (uint)_textRotation;
                cellFormat.ApplyAlignment = true;

            }
        }
        public Font Font
        {
            get
            {
                if(_font == null)
                    _font = new Font(Context.Styles.GetStyleFont(styleIndexCell.Value), this);

                return _font;
            }
        }
        public Border Borders
        {
            get
            {
                if (_borders == null)
                    _borders = new Border(Context.Styles.GetStyleBorder(styleIndexCell.Value), this);

                return _borders;
            }
        }
        public bool IsWrap
        {
            get
            {
                _isWrap = cellFormat?.Alignment?.WrapText?.Value ?? false;

                return _isWrap;
            }
            set
            {
                if (_isWrap == value) 
                    return;

                if (value)
                {
                    if (cellFormat.Alignment == null)
                        cellFormat.Alignment = new OpenXmlSpreadsheet.Alignment();

                    cellFormat.Alignment.WrapText = new OpenXml.BooleanValue(value);
                }
                else
                {
                    if (cellFormat.Alignment != null)
                    {
                        cellFormat.Alignment.WrapText = new OpenXml.BooleanValue(false);
                    }
                }
            }
        }

        internal StyleCell(Cell cell)
        {
            this._cell = cell;
        }
    }
}
