using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXMLight.config;
using DocumentFormat.OpenXml;

namespace OpenXMLight.Configurations.Formatting
{
    public readonly record struct HorizontalAlignments : IEnumValue<JustificationValues>
    {
        public JustificationValues Value => _value ?? JustificationValues.Left;


        public static HorizontalAlignments Left => new HorizontalAlignments(JustificationValues.Left);
        public static HorizontalAlignments Center => new HorizontalAlignments(JustificationValues.Center);
        public static HorizontalAlignments Right => new HorizontalAlignments(JustificationValues.Right);


        private readonly JustificationValues? _value;
        public HorizontalAlignments(JustificationValues jsValue)
        {
            _value = jsValue;
        }

        internal static HorizontalAlignments Parse(JustificationValues value)
        {
            return value switch
            {
                var v when v == JustificationValues.Left ||
                           v == JustificationValues.Start ||
                           v == JustificationValues.NumTab => HorizontalAlignments.Left,

                var v when v == JustificationValues.Center => HorizontalAlignments.Center,

                var v when v == JustificationValues.Right ||
                           v == JustificationValues.End => HorizontalAlignments.Right,
                _ => HorizontalAlignments.Left
            };
        }
    }
}
