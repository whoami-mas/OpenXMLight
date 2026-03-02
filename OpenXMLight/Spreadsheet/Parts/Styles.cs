using OpenXMLight.Configurations.Parts.InterfacesParts;
using OpenXMLight.Spreadsheet.Formatting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLight.Spreadsheet.Parts
{
    internal class Styles : IElementPart<OpenXmlPackaging.WorkbookStylesPart>
    {
        public OpenXmlPackaging.WorkbookStylesPart PartXml { get; set; }


        internal Styles(OpenXmlPackaging.WorkbookStylesPart stylePart)
        {
            PartXml = stylePart;

            CheckedExists();
        }


        public void CheckedExists()
        {
            PartXml.Stylesheet ??= new OpenXmlSpreadsheet.Stylesheet();

            GetFormatteId();
            AddStyle(TypeValue.General);
        }


        public TypeValue GetFormatteCell(int indexStyle)
        {
            if (PartXml.Stylesheet.CellFormats == null)
                throw new Exception("Пустой список форматов ячеек");

            var format = PartXml.Stylesheet.CellFormats.OfType<OpenXmlSpreadsheet.CellFormat>().ToList()[indexStyle];

            return TypeValue.Parse(format.NumberFormatId);
        }

        public uint AddStyle(TypeValue typeFormatte)
        {
            PartXml.Stylesheet.CellFormats ??= new OpenXmlSpreadsheet.CellFormats();

            OpenXmlSpreadsheet.CellFormat format = new()
            {
                NumberFormatId = GetNumberungFormatte(typeFormatte),
                FontId = GetFontId(),
                FillId = GetFillId(),
                BorderId = GetBorderId(),
                FormatId = GetFormatteId(),
                ApplyNumberFormat = GetNumberFormatte(typeFormatte)
            };

            PartXml.Stylesheet.CellFormats.AppendChild(format);

            return Convert.ToUInt32(PartXml.Stylesheet.CellFormats.OfType<OpenXmlSpreadsheet.CellFormat>().ToList().IndexOf(format));
        }
        public uint GetFormatteCellIndex(TypeValue typeFormatte)
        {
            if (PartXml.Stylesheet.CellFormats == null)
                return AddStyle(typeFormatte);

            OpenXmlSpreadsheet.CellFormat cellFormatte =
                PartXml.Stylesheet.CellFormats.OfType<OpenXmlSpreadsheet.CellFormat>().FirstOrDefault(f => f.NumberFormatId == typeFormatte.Value);

            if (cellFormatte == null)
                return AddStyle(typeFormatte);

            return Convert.ToUInt32(PartXml.Stylesheet.CellFormats.OfType<OpenXmlSpreadsheet.CellFormat>().ToList().IndexOf(cellFormatte));
        }


        private uint GetFontId()
        {
            PartXml.Stylesheet.Fonts ??= new OpenXmlSpreadsheet.Fonts();

            if(PartXml.Stylesheet.Fonts?.ChildElements.Count > 0)
                return 0;

            OpenXmlSpreadsheet.Font defaultFont = new OpenXmlSpreadsheet.Font()
            {
                FontSize = new OpenXmlSpreadsheet.FontSize() { Val = 11},
                FontName = new OpenXmlSpreadsheet.FontName() { Val = "Calibri"},
                FontFamilyNumbering = new OpenXmlSpreadsheet.FontFamilyNumbering() { Val = 2 },
                FontCharSet = new OpenXmlSpreadsheet.FontCharSet() { Val = 204 },
                FontScheme = new OpenXmlSpreadsheet.FontScheme() { Val = OpenXmlSpreadsheet.FontSchemeValues.Minor }
            };

            PartXml.Stylesheet.Fonts?.AppendChild(defaultFont);

            return 0;
        }
        private uint GetFillId()
        {
            PartXml.Stylesheet.Fills ??= new OpenXmlSpreadsheet.Fills();

            if (PartXml.Stylesheet.Fills?.ChildElements.Count > 0)
                return 0;

            OpenXmlSpreadsheet.Fill defaultFill = new OpenXmlSpreadsheet.Fill()
            {
                PatternFill = new OpenXmlSpreadsheet.PatternFill() { PatternType = OpenXmlSpreadsheet.PatternValues.None }
            };

            PartXml.Stylesheet.Fills?.AppendChild(defaultFill);

            return 0;
        }
        private uint GetBorderId()
        {
            PartXml.Stylesheet.Borders ??= new OpenXmlSpreadsheet.Borders();

            if (PartXml.Stylesheet.Borders?.ChildElements.Count > 0)
                return 0;

            OpenXmlSpreadsheet.Border defaultBorder = new OpenXmlSpreadsheet.Border()
            {
                LeftBorder = new OpenXmlSpreadsheet.LeftBorder(),
                RightBorder = new OpenXmlSpreadsheet.RightBorder(),
                TopBorder = new OpenXmlSpreadsheet.TopBorder(),
                BottomBorder = new OpenXmlSpreadsheet.BottomBorder(),
                DiagonalBorder = new OpenXmlSpreadsheet.DiagonalBorder()
            };

            PartXml.Stylesheet.Borders?.AppendChild(defaultBorder);

            return 0;
        }
        private uint GetFormatteId()
        {
            PartXml.Stylesheet.CellStyleFormats ??= new OpenXmlSpreadsheet.CellStyleFormats();

            if (PartXml.Stylesheet.CellStyleFormats?.ChildElements.Count > 0)
                return 0;

            OpenXmlSpreadsheet.CellFormat cellFormat = new OpenXmlSpreadsheet.CellFormat()
            {
                NumberFormatId = 0,
                FontId = GetFontId(),
                FillId = GetFillId(),
                BorderId = GetBorderId(),
            };

            PartXml.Stylesheet.CellStyleFormats?.AppendChild(cellFormat);
            //PartXml.Stylesheet.CellStyleFormats?.Count = ;

            return 0;
        }
        

        private bool GetNumberFormatte(TypeValue typeCell)
        {
            bool result = false;

            if (typeCell == TypeValue.Date
                ||
                typeCell == TypeValue.Number
                ||
                typeCell == TypeValue.Percent)
                result = true;


            return result;
        }
        private uint GetNumberungFormatte(TypeValue typeCell)
        {
            PartXml.Stylesheet.NumberingFormats ??= new OpenXmlSpreadsheet.NumberingFormats();

            switch (typeCell)
            {
                case var type when type == TypeValue.Date:
                    OpenXmlSpreadsheet.NumberingFormat? nmbFormat = PartXml.Stylesheet.NumberingFormats.Elements<OpenXmlSpreadsheet.NumberingFormat>().FirstOrDefault(f => f.NumberFormatId == typeCell.Value);
                    if(nmbFormat == null)
                    {
                        nmbFormat = new OpenXmlSpreadsheet.NumberingFormat()
                        {
                            NumberFormatId = Convert.ToUInt32(typeCell.Value),
                            FormatCode = "dd.mm.yyyy"
                        };

                        PartXml.Stylesheet.NumberingFormats.AppendChild(nmbFormat);
                    }
                    break;
            }

            return Convert.ToUInt32(typeCell.Value);
        }
    }
}
