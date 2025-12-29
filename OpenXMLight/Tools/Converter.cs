using OpenXMLight.Configurations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Tools
{
    internal static class Converter
    {
        internal static string ConvertWidthToCm(string width) => (double.Parse(width) / Configuration.WidthCm).ToString();
        internal static string ConvertWidthToPercent(string width) => (double.Parse(width) / 50).ToString();
        internal static string ConvertWidthPercentToTwips(string width) => (double.Parse(width) * Configuration.WidthCm).ToString();
        internal static string ConvertWidthCmToTwips(string width) => (double.Parse(width) * 50).ToString();
    }
}
