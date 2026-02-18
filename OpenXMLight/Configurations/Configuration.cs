using OpenXMLight.Configurations.Elements;
using OpenXMLight.Configurations.Elements.TableElements;
using OpenXMLight.Configurations.Elements.TableElements.Formattings;
using OpenXMLight.Configurations.Elements.TableElements.Formattings.WidthComponents;
using OpenXMLight.Configurations.Formatting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Configurations
{
    internal static class Configuration
    {
        internal static int LineWidthInTable => 8;
        internal static int TwipsInPixels => 15;
        internal static int InchInPixels => 20;
        internal static int WidthCm => 567;
        internal static int DEFAULT_FONTSIZE => 11;
        internal static FontsFamily DEFAULT_FONTFAMILY => FontsFamily.Arial;
        internal static HorizontalAlignments DEFAULT_HORIZONTAL_ALIGNMENT => HorizontalAlignments.Left;
        internal static SpacingBetweenLines DEFAULT_SPACING_BETWEEN_LINES => new SpacingBetweenLines();
        internal static BordersLine? DEFAULT_BORDERS_LINE => null;
        internal static TableCellWidth<TableWidth> DEFAULT_TABLEWIDTH => new TableCellWidth<TableWidth>();
        internal static TableCellWidth<CellWidth> DEFAULT_CELLWIDTH => new TableCellWidth<CellWidth>();
        internal static VerticalAlignments DEFAULT_VERTICAL_ALIGNMENT => VerticalAlignments.Top;
        internal static TextDirectionType? DEFAULT_TEXTDIRECTION => null;
        internal static Color? DEFAULT_COLOR_TEXT => null;
        internal static Color? DEFAULT_COLOR_SHADE => null;
    }
}
