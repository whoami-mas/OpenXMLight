using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXMLight.config;

namespace OpenXMLight.Spreadsheet.Formatting
{
    public readonly record struct TypeValue : IEnumValue<string>
    {
        public string Value => _value;
        

        public static TypeValue General => new TypeValue("General");
        public static TypeValue Percent => new TypeValue("0%");


        private readonly string _value = "General";


        public TypeValue(string value)
        {
            _value = value;
        }
    }
}
