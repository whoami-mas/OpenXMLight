using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXMLight.Configurations.Formatting;

using OpenXml = DocumentFormat.OpenXml;

namespace OpenXMLight.Spreadsheet.Formatting
{
    public interface IBorder
    {
        public BorderStyle? Type { get; set; }
        public Color? Color { get; set; }
    }
}
