using OpenXMLight.Configurations;
using OpenXMLight.Configurations.Elements;
using OpenXMLight.Configurations.Elements.TableElements;
using OpenXMLight.Configurations.Elements.TableElements.Formattings;
using OpenXMLight.Configurations.Elements.TableElements.Formattings.WidthComponents;
using OpenXMLight.Configurations.Formatting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using OpenXML = DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlFormatte = DocumentFormat.OpenXml;

namespace OpenXMLight.Tools
{
    internal static class HelperData
    {
        internal static int GetRowIndex(string input)
        {
            Match regexMatch = Regex.Match(input, @"\d+", RegexOptions.IgnoreCase);
            if (regexMatch.Success)
            {
                return int.Parse(regexMatch.Value);
            }
            else
                return 0 ;
        }
        internal static int GetColumnIndex(string input)
        {
            int index = 0;
            string column = Regex.Match(input, @"[A-Z]+", RegexOptions.IgnoreCase).Value;

            for (int i = 0; i < column.Length; i++)
            {
                index *= 26;
                index += (column[i] - 'A' + 1);
            }
            return index;
        }
        internal static string GetColumnByIndex(int index)
        {
            string columnName = string.Empty;
            while (index > 0)
            {
                int remainder = (index - 1) % 26;
                columnName = (char)(remainder + 'A') + columnName;
                index = (index - 1) / 26;
            }
            return columnName;
        }


        internal static bool TryParseFontSize(string value, out int fontSize)
        {
            fontSize = Configuration.DEFAULT_FONTSIZE;

            if (string.IsNullOrWhiteSpace(value))
                return false;

            if (!int.TryParse(value, out int halfPoints) || halfPoints < 0)
                return false;

            fontSize = halfPoints / 2;
            return true;
        }
        internal static bool TryParseFontFamily(string value, out FontsFamily fontFamily)
        {
            fontFamily = Configuration.DEFAULT_FONTFAMILY;

            if (string.IsNullOrWhiteSpace(value))
                return false;

            fontFamily = FontsFamily.Parse(value);

            return true;
        }
        internal static bool TryParseParagraphAlignment(object? value, out HorizontalAlignments alignment)
        {
            alignment = Configuration.DEFAULT_HORIZONTAL_ALIGNMENT;

            if (value is not OpenXML.JustificationValues justification)
                return false;

            alignment = HorizontalAlignments.Parse(justification);

            return true;
        }
        internal static bool TryParseSpacingBetweenLines(OpenXML.SpacingBetweenLines? spacingXml, out SpacingBetweenLines spacing)
        {
            spacing = Configuration.DEFAULT_SPACING_BETWEEN_LINES;

            if (spacingXml == null)
                return false;

            if (!int.TryParse(spacingXml.After, out int after)
                                    ||
                !int.TryParse(spacingXml.Before, out int before)
                                    ||
                !int.TryParse(spacingXml.Line, out int line))
                return false;

            spacing.After = after;
            spacing.Before = before;
            spacing.Line = line;

            return true;
        }
        internal static bool TryParseTableBorders(OpenXML.TableBorders? tblBorders, out BordersLine? borders)
        {
            borders = Configuration.DEFAULT_BORDERS_LINE;

            if (tblBorders == null)
                return false;

            borders = new(tblBorders);
            
            return true;
        }
        internal static bool TryParseTableWidth(OpenXML.TableWidth? tblWidth, out TableCellWidth<TableWidth> width)
        {
            width = Configuration.DEFAULT_TABLEWIDTH;

            if (tblWidth == null)
                return false;

            width = new(tblWidth);

            return true;
        }
        internal static bool TryParseCellWidth(OpenXML.TableCellWidth? cellWidth, out TableCellWidth<CellWidth> width)
        {
            width = Configuration.DEFAULT_CELLWIDTH;

            if (cellWidth == null)
                return false;

            width = new(cellWidth);

            return true;
        }
        internal static bool TryParseTableCellVerticalAlignment(object? cellAlignment, out VerticalAlignments alignment)
        {
            alignment = Configuration.DEFAULT_VERTICAL_ALIGNMENT;

            if (cellAlignment is not OpenXML.TableVerticalAlignmentValues cellVerticalAlignment)
                return false;

            alignment = VerticalAlignments.Parse(cellVerticalAlignment);

            return true;
        }
        internal static bool TryParseTextDirectionCell(object? textDir, out TextDirectionType? textDirection)
        {
            textDirection = Configuration.DEFAULT_TEXTDIRECTION;

            if (textDir is not OpenXML.TextDirectionValues _textDirection)
                return false;

            textDirection = TextDirectionType.Parse(_textDirection);

            return true;
        }
        internal static bool TryParseColorText(object? colorXml, out Color? color)
        {
            color = Configuration.DEFAULT_COLOR_TEXT;

            if (colorXml is not OpenXmlFormatte.StringValue hex)
                return false;

            color = Color.FromHex(hex.ToString().Insert(0, "#"));
            
            return true;
        }
        internal static bool TryParseColorShade(object? shdCellXml, out Color? color)
        {
            color = Configuration.DEFAULT_COLOR_SHADE;

            if (shdCellXml is not OpenXML.Shading shd)
                return false;

            color = Color.FromHex(shd.Fill);

            return true;
        }
    }
}
