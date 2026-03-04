using OpenXMLight.Configurations.Formatting;
using OpenXMLight.Spreadsheet.ExcelContext;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using OpenXml = DocumentFormat.OpenXml;
using OpenXMLight.Spreadsheet.Elements;

namespace OpenXMLight.Spreadsheet.Formatting
{
    public class Font
    {
        private Context Context => Context.Instance;
        internal OpenXmlSpreadsheet.Font? fontXml;
        internal StyleCell style;

        private int _size;
        private bool _bold;
        private FontsFamily _fontFamily;
        
        public int Size
        {
            get 
            {
                _size = Convert.ToInt32(fontXml.FontSize.Val.Value);

                return _size;
            }
            set
            {
                _size = value;
            }
        }
        public bool Bold
        {
            get
            {
                _bold = fontXml.Bold != null
                    ? true
                    : false;

                return _bold;
            }
            set
            {
                _bold = value;
            }
        }
        public FontsFamily FontFamily
        {
            get
            {
                _fontFamily = FontsFamily.Parse(fontXml.FontName.Val.Value);
                
                return _fontFamily;
            }
            set
            {
                _fontFamily = value;
            }
        }


        internal Font(OpenXmlSpreadsheet.Font fontXml, StyleCell style)
        {
            this.fontXml = fontXml;
            this.style = style;
        }

        public void CommitChange()
        {
            OpenXmlSpreadsheet.Font fontNew = new()
            {
                FontSize = new OpenXmlSpreadsheet.FontSize() { Val = _size },
                FontName = new OpenXmlSpreadsheet.FontName() { Val = _fontFamily.Value },
                Bold = _bold ? new OpenXmlSpreadsheet.Bold() : null
            };
            
            Context.Styles.CheckStyleFont(ref style, ref fontNew);

            fontXml = fontNew;
        }
    }
}
