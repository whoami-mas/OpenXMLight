using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using OpenXml = DocumentFormat.OpenXml;
using Format = OpenXMLight.Configurations.Formatting;
using OpenXMLight.Tools.Colors;

namespace OpenXMLight.Spreadsheet.Formatting
{
    public class RightBorder : IBorder
    {
        internal OpenXmlSpreadsheet.RightBorder border { get; }

        private BorderStyle? _type;
        private Format.Color? _color;

        public BorderStyle? Type
        {
            get
            {
                if(_type == null)
                    if (border.Style != null && !border.Style.HasValue)
                        _type = BorderStyle.Parse(border.Style);

                return _type;
            }
            set
            {
                _type = value;
            }
        }
        public Format.Color? Color
        {
            get
            {
                if (border.Color != null)
                    if (border.Color.Indexed != null)
                        _color = ColorsIndex.indexedColor[border.Color.Indexed];
                    else if (border.Color.Rgb != null)
                        _color = Format.Color.FromHex(border.Color.Rgb.Value);

                return _color;
            }
            set
            {
                _color = value;
            }
        }



        internal RightBorder(OpenXmlSpreadsheet.RightBorder rightBorder)
        {
            this.border = rightBorder;
        }
    }
}
