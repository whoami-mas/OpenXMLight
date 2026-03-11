using OpenXMLight.Configurations.Parts.InterfacesParts;
using OpenXMLight.Spreadsheet.Elements;
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

            GetDefaultFormatteId();
            AddStyle(TypeValue.General);
        }

        public uint GetFormatteCellIndex(TypeValue typeFormatte)
        {
            if (PartXml.Stylesheet.CellFormats == null || PartXml.Stylesheet.CellFormats.ChildElements.Count <= 1)
                return AddStyle(typeFormatte);

            OpenXmlSpreadsheet.CellFormat cellFormatte =
                PartXml.Stylesheet.CellFormats.OfType<OpenXmlSpreadsheet.CellFormat>().FirstOrDefault(f => f.NumberFormatId == typeFormatte.Value);

            if (cellFormatte == null)
                return AddStyle(typeFormatte);

            return Convert.ToUInt32(PartXml.Stylesheet.CellFormats.OfType<OpenXmlSpreadsheet.CellFormat>().ToList().IndexOf(cellFormatte));
        }



        private void ExistsFormatte()
        {
            if (PartXml.Stylesheet.CellFormats == null)
                throw new Exception("Пустой список форматов ячеек");
        }

        public OpenXmlSpreadsheet.CellFormat? GetFormatteCell(uint indexStyle)
        {
            ExistsFormatte();

            return PartXml.Stylesheet.CellFormats.OfType<OpenXmlSpreadsheet.CellFormat>().ToList()[(int)indexStyle];
        }
        public OpenXmlSpreadsheet.Font GetStyleFont(uint indexStyle)
        {
            OpenXmlSpreadsheet.CellFormat format = GetFormatteCell(indexStyle);

            return PartXml.Stylesheet.Fonts.Elements<OpenXmlSpreadsheet.Font>().ToList()[(int)format.FontId.Value];
        }
        public OpenXmlSpreadsheet.Border GetStyleBorder(uint indexStyle)
        {
            OpenXmlSpreadsheet.CellFormat format = GetFormatteCell(indexStyle);

            return PartXml.Stylesheet.Borders.Elements<OpenXmlSpreadsheet.Border>().ToList()[(int)format.BorderId.Value];
        }



        public uint AddStyle(TypeValue typeFormatte)
        {
            PartXml.Stylesheet.CellFormats ??= new OpenXmlSpreadsheet.CellFormats();

            OpenXmlSpreadsheet.CellFormat format = new()
            {
                NumberFormatId = GetNumberungFormatte(typeFormatte),
                FontId = GetDefaultFontId(),
                FillId = GetDefaultFillId(),
                BorderId = GetDefaultBorderId(),
                FormatId = GetDefaultFormatteId(),
                ApplyNumberFormat = GetNumberFormatte(typeFormatte) ? true : null
            };
            
            PartXml.Stylesheet.CellFormats.AppendChild(format);

            return Convert.ToUInt32(PartXml.Stylesheet.CellFormats.OfType<OpenXmlSpreadsheet.CellFormat>().ToList().IndexOf(format));
        }
        public OpenXmlSpreadsheet.CellFormat AddStyle(uint styleIndex)
        {
            OpenXmlSpreadsheet.CellFormat? cellFormat = PartXml.Stylesheet.CellFormats.ToList()[(int)styleIndex] as OpenXmlSpreadsheet.CellFormat;
            
            return (OpenXmlSpreadsheet.CellFormat)cellFormat.CloneNode(true);
        }
        public uint AddStyle(OpenXmlSpreadsheet.CellFormat template)
        {
            OpenXmlSpreadsheet.CellFormat cellFormat = (OpenXmlSpreadsheet.CellFormat)template.CloneNode(true);
            PartXml.Stylesheet.CellFormats.AppendChild(cellFormat);


            return Convert.ToUInt32(PartXml.Stylesheet.CellFormats.OfType<OpenXmlSpreadsheet.CellFormat>().ToList().IndexOf(cellFormat));
        }



        public void CheckStyleFont(ref StyleCell style, ref OpenXmlSpreadsheet.Font fontNew)
        {
            var fonts = PartXml.Stylesheet.Fonts;
            OpenXmlSpreadsheet.CellFormat format = AddStyle(style.styleIndexCell.Value);

            int hashFontNew = GetHashCodeFont(fontNew);
            for(int i = 0; i < fonts.ChildElements.Count; i++)
            {
                if (fonts.ChildElements[i] is OpenXmlSpreadsheet.Font font && 
                    GetHashCodeFont(font) == hashFontNew)
                {
                    format.FontId = (uint)i;

                    format = CheckFormatteCell(ref format, ref style);
                    return;
                }
            }

            fonts.AppendChild(fontNew);
            format.FontId = Convert.ToUInt32(fonts.ToList().IndexOf(fontNew));
            format = CheckFormatteCell(ref format, ref style);
        }
        public void CheckStyleBorder(ref StyleCell style, ref OpenXmlSpreadsheet.Border borderNew)
        {
            var borders = PartXml.Stylesheet.Borders;
            OpenXmlSpreadsheet.CellFormat format = AddStyle(style.styleIndexCell.Value);

            int hashBorderNew = GetHashCodeBorder(borderNew);
            for(int i = 0; i < borders.ChildElements.Count; i++)
            {
                if (borders.ChildElements[i] is OpenXmlSpreadsheet.Border border &&
                    GetHashCodeBorder(border) == hashBorderNew)
                {
                    format.BorderId = (uint)i;

                    format.ApplyBorder = true;

                    format = CheckFormatteCell(ref format, ref style);
                    return;
                }
            }

            borders.AppendChild(borderNew);
            format.BorderId = Convert.ToUInt32(borders.ToList().IndexOf(borderNew));
            format.ApplyBorder = true;
            format = CheckFormatteCell(ref format, ref style);
        }      
        private OpenXmlSpreadsheet.CellFormat CheckFormatteCell(ref OpenXmlSpreadsheet.CellFormat format, ref StyleCell style)
        {
            var cellFormatts = PartXml.Stylesheet.CellFormats;

            int hashCellFormat = GetHashCodeCellFormatte(format);
            for(int i = 0; i < cellFormatts.ChildElements.Count; i++)
            {
                if (cellFormatts.ChildElements[i] is OpenXmlSpreadsheet.CellFormat formatOld &&
                    GetHashCodeCellFormatte(formatOld) == hashCellFormat)
                {
                    if(GetHashCodeAlignment(formatOld.Alignment) == GetHashCodeAlignment(format.Alignment))
                    {
                        style.cellXml.StyleIndex = (uint)i;
                        return formatOld;
                    }

                    continue;
                }
            }

            cellFormatts.AppendChild(format);
            style.cellXml.StyleIndex = (uint)cellFormatts.ToList().IndexOf(format);
            return format;
        }
        
        
        public uint IsFirstFormatteCell(uint indexStyle)
        {
            if (indexStyle != 0)
                return indexStyle;

            return AddStyle(GetFormatteCell(indexStyle));
        }

        #region Hash code
        private int GetHashCodeFont(OpenXmlSpreadsheet.Font font)
        {
            unchecked
            {
                int hash = 17;

                hash = hash * 23 + font.FontSize.Val.GetHashCode();
                hash = hash * 23 + (font.Bold != null ? 1 : 0);
                hash = hash * 23 + font.FontName.Val.GetHashCode();

                return hash;
            }
        }
        private int GetHashCodeBorder(OpenXmlSpreadsheet.Border border)
        {
            unchecked
            {
                int hash = 17;

                //Left
                hash = hash * 23 + (border.LeftBorder?.Style?.HasValue == true ? border.LeftBorder.Style.GetHashCode() : 0);
                hash = hash * 23 + GetHashCodeBorderColor(border.LeftBorder?.Color);

                //Right
                hash = hash * 23 + (border.RightBorder?.Style?.HasValue == true ? border.RightBorder.Style.GetHashCode() : 0);
                hash = hash * 23 + GetHashCodeBorderColor(border.RightBorder?.Color);

                //Top
                hash = hash * 23 + (border.TopBorder?.Style?.HasValue == true ? border.TopBorder.Style.GetHashCode() : 0);
                hash = hash * 23 + GetHashCodeBorderColor(border.TopBorder?.Color);

                //Bottom
                hash = hash * 23 + (border.BottomBorder?.Style?.HasValue == true ? border.BottomBorder.Style.GetHashCode() : 0);
                hash = hash * 23 + GetHashCodeBorderColor(border.BottomBorder?.Color);

                return hash;
            }
        }
        private int GetHashCodeCellFormatte(OpenXmlSpreadsheet.CellFormat format)
        {
            unchecked
            {
                int hash = 17;

                hash = hash * 23 + format.NumberFormatId.GetHashCode();
                hash = hash * 23 + format.FontId.GetHashCode();
                hash = hash * 23 + format.FillId.GetHashCode();
                hash = hash * 23 + format.BorderId.GetHashCode();
                hash = hash * 23 + format.FormatId.GetHashCode();

                return hash;
            }
        }
        private int GetHashCodeAlignment(OpenXmlSpreadsheet.Alignment alignment)
        {
            if (alignment == null)
                return 0;

            unchecked
            {
                int hash = 17;

                hash = hash * 23 + (alignment.TextRotation?.HasValue == true ? alignment.TextRotation.Value.GetHashCode() : 0);
                hash = hash * 23 + (alignment.Vertical?.HasValue == true ? alignment.Vertical.Value.GetHashCode() : 0);
                hash = hash * 23 + (alignment.Horizontal?.HasValue == true ? alignment.Horizontal.Value.GetHashCode() : 0);

                return hash;
            }
        }

        private int GetHashCodeBorderColor(OpenXmlSpreadsheet.Color? color)
        {
            unchecked
            {
                if (color == null)
                    return 0;

                if (color.Rgb != null && !string.IsNullOrWhiteSpace(color.Rgb.Value))
                    return color.Rgb.Value.GetHashCode();

                if (!color.Indexed.HasValue)
                    return color.Indexed.Value.GetHashCode();

                return 0;
            }
        }
        #endregion

        #region Default

        private uint GetDefaultFontId()
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
        private uint GetDefaultFillId()
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
        private uint GetDefaultBorderId()
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
        private uint GetDefaultFormatteId()
        {
            PartXml.Stylesheet.CellStyleFormats ??= new OpenXmlSpreadsheet.CellStyleFormats();

            if (PartXml.Stylesheet.CellStyleFormats?.ChildElements.Count > 0)
                return 0;

            OpenXmlSpreadsheet.CellFormat cellFormat = new OpenXmlSpreadsheet.CellFormat()
            {
                NumberFormatId = 0,
                FontId = GetDefaultFontId(),
                FillId = GetDefaultFillId(),
                BorderId = GetDefaultBorderId(),
            };

            PartXml.Stylesheet.CellStyleFormats?.AppendChild(cellFormat);
            //PartXml.Stylesheet.CellStyleFormats?.Count = ;

            return 0;
        }

        #endregion

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
