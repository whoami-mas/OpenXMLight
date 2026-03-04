using OpenXMLight.config;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLight.Spreadsheet.Formatting
{
    public readonly record struct VerticalAlignments : IEnumValue<OpenXmlSpreadsheet.VerticalAlignmentValues>
    {
        public OpenXmlSpreadsheet.VerticalAlignmentValues Value => _value;


        public static VerticalAlignments Top => new VerticalAlignments(OpenXmlSpreadsheet.VerticalAlignmentValues.Top);
        public static VerticalAlignments Center => new VerticalAlignments(OpenXmlSpreadsheet.VerticalAlignmentValues.Center);
        public static VerticalAlignments Bottom => new VerticalAlignments(OpenXmlSpreadsheet.VerticalAlignmentValues.Bottom);



        private readonly OpenXmlSpreadsheet.VerticalAlignmentValues _value;


        public VerticalAlignments(OpenXmlSpreadsheet.VerticalAlignmentValues value)
        {
            _value = value;
        }


        public static VerticalAlignments Parse(OpenXmlSpreadsheet.VerticalAlignmentValues value)
        {
            if (value == null)
                return VerticalAlignments.Top;

            return value switch
            {
                var v when v == OpenXmlSpreadsheet.VerticalAlignmentValues.Top => VerticalAlignments.Top,
                var v when v == OpenXmlSpreadsheet.VerticalAlignmentValues.Center => VerticalAlignments.Center,
                var v when v == OpenXmlSpreadsheet.VerticalAlignmentValues.Bottom => VerticalAlignments.Bottom,
                _ => VerticalAlignments.Top,
            };
        }
    }
}
