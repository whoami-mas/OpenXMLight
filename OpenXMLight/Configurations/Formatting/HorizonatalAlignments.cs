using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXMLight.config;

namespace OpenXMLight.Configurations.Formatting
{
    public readonly record struct HorizonatalAlignments : IEnumValue<JustificationValues>
    {
        public JustificationValues Value => _value ?? JustificationValues.Left;

        public static HorizonatalAlignments Left => new HorizonatalAlignments(JustificationValues.Left);
        public static HorizonatalAlignments Center => new HorizonatalAlignments(JustificationValues.Center);
        public static HorizonatalAlignments Right => new HorizonatalAlignments(JustificationValues.Right);


        private readonly JustificationValues? _value;
        public HorizonatalAlignments(JustificationValues jsValue)
        {
            _value = jsValue;
        }
    }
}
