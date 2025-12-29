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
        internal static int LineWidthInTable { get; set; } = 8;
        internal static int TwipsInPixels { get; set; } = 15;
        internal static int InchInPixels { get; set; } = 20;
        internal static int WidthCm { get; set; } = 567;
        internal static int DEFAULT_FONTSIZE { get; set; } = 11;
        internal static FontsFamily DEFAULT_FONTFAMILY { get; set; } = FontsFamily.Arial;
        internal static HorizontalAlignments DEFAULT_HORIZONTAL_ALIGNMENT { get; set; } = HorizontalAlignments.Left;
        internal static SpacingBetweenLines DEFAULT_SPACING_BETWEEN_LINES { get; set; } = new SpacingBetweenLines();
        internal static BordersLine? DEFAULT_BORDERS_LINE { get; set; } = null;
        internal static TableCellWidth<TableWidth> DEFAULT_TABLEWIDTH { get; set; } = new TableCellWidth<TableWidth>();
        internal static TableCellWidth<CellWidth> DEFAULT_CELLWIDTH { get; set; } = new TableCellWidth<CellWidth>();
        internal static VerticalAlignments DEFAULT_VERTICAL_ALIGNMENT { get; set; } = VerticalAlignments.Top;
    }
}
