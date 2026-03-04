using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using OpenXml = DocumentFormat.OpenXml;
using OpenXMLight.config;

namespace OpenXMLight.Spreadsheet.Formatting
{
    public readonly record struct BorderStyle : IEnumValue<OpenXmlSpreadsheet.BorderStyleValues>
    {
        public OpenXmlSpreadsheet.BorderStyleValues Value => _value;



        public static BorderStyle Single => new BorderStyle(OpenXmlSpreadsheet.BorderStyleValues.Thin);
        public static BorderStyle Double => new BorderStyle(OpenXmlSpreadsheet.BorderStyleValues.Double);
        public static BorderStyle Dashed => new BorderStyle(OpenXmlSpreadsheet.BorderStyleValues.Dashed);
        public static BorderStyle None => new BorderStyle(OpenXmlSpreadsheet.BorderStyleValues.None);




        private readonly OpenXmlSpreadsheet.BorderStyleValues _value = OpenXmlSpreadsheet.BorderStyleValues.Thin;
        public BorderStyle(OpenXmlSpreadsheet.BorderStyleValues value)
        {
            _value = value;
        }

        public static BorderStyle Parse(OpenXmlSpreadsheet.BorderStyleValues value)
        {
            return value switch
            {
                var v when v == OpenXmlSpreadsheet.BorderStyleValues.Thin => BorderStyle.Single,
                var v when v == OpenXmlSpreadsheet.BorderStyleValues.Double => BorderStyle.Double,
                var v when v == OpenXmlSpreadsheet.BorderStyleValues.Dashed => BorderStyle.Dashed,
                var v when v == OpenXmlSpreadsheet.BorderStyleValues.None => BorderStyle.None,
            };
        }
    }
}
