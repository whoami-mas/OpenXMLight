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
        private Context Context => Context.Instance;


        internal OpenXmlSpreadsheet.Cell? cellXml;
        internal uint? styleIndexCell => cellXml?.StyleIndex;


        private TypeValue _typeValue = TypeValue.General;
        private HorizontalAlignments? _hAlignment;
        private Font _font;
        private Border _borders;

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
                    var formatte = Context.Styles.GetFormatteCell(styleIndexCell.Value);

                    _hAlignment = HorizontalAlignments.Parse(formatte.Alignment?.Horizontal);
                }

                return _hAlignment;
            }
            set
            {
                var formatte = Context.Styles.GetFormatteCell(styleIndexCell.Value);
                
                formatte.Alignment ??= new OpenXmlSpreadsheet.Alignment();
                formatte.Alignment.Horizontal = value.Value.Value;
                formatte.ApplyAlignment = true;


                _hAlignment = value;
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


        internal StyleCell(OpenXmlSpreadsheet.Cell? cellXml)
        {
            this.cellXml = cellXml;
        }
    }
}
