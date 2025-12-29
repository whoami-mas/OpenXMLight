using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLight.config;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Configurations.Formatting
{
    public readonly record struct BordersType : IEnumValue<BorderValues>
    {
        public BorderValues Value => _value ?? BorderValues.Single;
        

        public static BordersType Single => new BordersType(BorderValues.Single);
        public static BordersType None => new BordersType(BorderValues.None);
        public static BordersType Double => new BordersType(BorderValues.Double);
        public static BordersType Dotted => new BordersType(BorderValues.Dotted);
        public static BordersType Dashed => new BordersType(BorderValues.Dashed);
        public static BordersType DotDash => new BordersType(BorderValues.DotDash);
        public static BordersType Triple => new BordersType(BorderValues.Triple);


        private readonly BorderValues? _value;
        public BordersType(BorderValues value)
        {
            _value = value;
        }


        internal static BordersType Parse(BorderValues? value)
        {
            if(!value.HasValue)
                return BordersType.Single;

            return value switch
            {
                var v when v == BorderValues.Single => BordersType.Single,
                var v when v == BorderValues.None => BordersType.None,
                var v when v == BorderValues.Double => BordersType.Double,
                var v when v == BorderValues.Dotted => BordersType.Dotted,
                var v when v == BorderValues.Dashed => BordersType.Dashed,
                var v when v == BorderValues.DotDash => BordersType.DotDash,
                var v when v == BorderValues.Triple => BordersType.Triple,
                _ => throw new ArgumentNullException("NULL")
            };
        }
    }
}
