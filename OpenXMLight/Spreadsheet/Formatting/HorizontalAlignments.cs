using OpenXMLight.config;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLight.Spreadsheet.Formatting
{
    public readonly record struct HorizontalAlignments : IEnumValue<OpenXmlSpreadsheet.HorizontalAlignmentValues>
    {
        public OpenXmlSpreadsheet.HorizontalAlignmentValues Value => _value;


        public static HorizontalAlignments Left => new HorizontalAlignments(OpenXmlSpreadsheet.HorizontalAlignmentValues.Left);
        public static HorizontalAlignments Center => new HorizontalAlignments(OpenXmlSpreadsheet.HorizontalAlignmentValues.Center);
        public static HorizontalAlignments Right => new HorizontalAlignments(OpenXmlSpreadsheet.HorizontalAlignmentValues.Right);



        private readonly OpenXmlSpreadsheet.HorizontalAlignmentValues _value;


        public HorizontalAlignments(OpenXmlSpreadsheet.HorizontalAlignmentValues value)
        {
            _value = value;
        }


        public static HorizontalAlignments Parse(OpenXmlSpreadsheet.HorizontalAlignmentValues value)
        {
            if (value == null)
                return HorizontalAlignments.Left;

            return value switch
            {
                var v when v == OpenXmlSpreadsheet.HorizontalAlignmentValues.Left => HorizontalAlignments.Left,
                var v when v == OpenXmlSpreadsheet.HorizontalAlignmentValues.Center => HorizontalAlignments.Center,
                var v when v == OpenXmlSpreadsheet.HorizontalAlignmentValues.Right => HorizontalAlignments.Right,
                _ => HorizontalAlignments.Left,
            };
        }
    }
}
