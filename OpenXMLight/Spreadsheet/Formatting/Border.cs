using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using OpenXml = DocumentFormat.OpenXml;
using OpenXMLight.Spreadsheet.Elements;
using OpenXMLight.Spreadsheet.ExcelContext;

namespace OpenXMLight.Spreadsheet.Formatting
{
    public class Border
    {
        private Context Context => Context.Instance;
        internal OpenXmlSpreadsheet.Border? borderXml;
        internal StyleCell style;


        private TopBorder _top;
        private LeftBorder _left;
        private RightBorder _right;
        private BottomBorder _bottom;


        public TopBorder Top
        {
            get => _top;
        }
        public LeftBorder Left
        {
            get => _left;
        }
        public RightBorder Right
        {
            get => _right;
        }
        public BottomBorder Bottom
        {
            get => _bottom;
        }

        internal Border(OpenXmlSpreadsheet.Border borderXml, StyleCell style)
        {
            this.borderXml = borderXml;
            this.style = style;

            this._top = new TopBorder(this.borderXml.TopBorder);
            this._left = new LeftBorder(this.borderXml.LeftBorder);
            this._right = new RightBorder(this.borderXml.RightBorder);
            this._bottom = new BottomBorder(this.borderXml.BottomBorder);
        }

        public void CommitChange()
        {
            OpenXmlSpreadsheet.Border borderNew = new(
                new OpenXmlSpreadsheet.LeftBorder()
                {
                    Style = Left.Type?.Value,
                    Color = (Left.Color.HasValue ?
                    new OpenXmlSpreadsheet.Color()
                    {
                        Rgb = new OpenXml.HexBinaryValue()
                        {
                            Value = Left.Color?.RGB
                        }
                    }
                    : null
                    )
                },
                new OpenXmlSpreadsheet.RightBorder()
                {
                    Style = Right.Type?.Value,
                    Color = (Right.Color.HasValue ?
                    new OpenXmlSpreadsheet.Color()
                    {
                        Rgb = new OpenXml.HexBinaryValue()
                        {
                            Value = Right.Color?.RGB
                        }
                    }
                    : null
                    )
                },
                new OpenXmlSpreadsheet.TopBorder()
                {
                    Style = Top.Type?.Value,
                    Color = (Top.Color.HasValue ?
                    new OpenXmlSpreadsheet.Color()
                    {
                        Rgb = new OpenXml.HexBinaryValue()
                        {
                            Value = Top.Color?.RGB
                        }
                    }
                    : null
                    )
                },
                new OpenXmlSpreadsheet.BottomBorder()
                {
                    Style = Bottom.Type?.Value,
                    Color = (Bottom.Color.HasValue ? 
                    new OpenXmlSpreadsheet.Color()
                    {
                        Rgb = new OpenXml.HexBinaryValue()
                        {
                            Value = Bottom.Color?.RGB
                        }
                    }
                    : null
                    )
                }
            );

            Context.Styles.CheckStyleBorder(ref style, ref borderNew);

            borderXml = borderNew;
        }
    }
}
