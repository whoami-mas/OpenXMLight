using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLight.config;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Configurations.Formatting
{
    public readonly record struct TextDirectionType : IEnumValue<TextDirectionValues?>
    {
        public TextDirectionValues? Value => _value;


        public static TextDirectionType LeftToRightTopToBottom => new TextDirectionType(TextDirectionValues.LefToRightTopToBottom);
        public static TextDirectionType TopToBottomRightToLeft => new TextDirectionType(TextDirectionValues.TopToBottomRightToLeft);
        public static TextDirectionType BottomToTopLeftToRight => new TextDirectionType (TextDirectionValues.BottomToTopLeftToRight);


        private readonly TextDirectionValues? _value;
        public TextDirectionType(TextDirectionValues value)
        {
            this._value = value;
        }

        internal static TextDirectionType? Parse(TextDirectionValues? value)
        {
            if (!value.HasValue)
                return null;

            return value switch
            {
                var v when v == TextDirectionValues.LefToRightTopToBottom => TextDirectionType.LeftToRightTopToBottom,
                var v when v == TextDirectionValues.TopToBottomRightToLeft => TextDirectionType.TopToBottomRightToLeft,
                var v when v == TextDirectionValues.BottomToTopLeftToRight => TextDirectionType.BottomToTopLeftToRight,
                _ => null
            };
        }
    }
}
