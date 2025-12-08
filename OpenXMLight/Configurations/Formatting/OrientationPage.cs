using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLight.config;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Configurations.Formatting
{
    public readonly record struct OrientationPage : IEnumValue<PageOrientationValues>
    {
        public PageOrientationValues Value => _value ?? PageOrientationValues.Portrait;


        public static OrientationPage Landscape => new OrientationPage(PageOrientationValues.Landscape);
        public static OrientationPage Portrait => new OrientationPage(PageOrientationValues.Portrait);


        private static PageOrientationValues? _value;
        public OrientationPage(PageOrientationValues value)
        {
            _value = value;
        }
    }
}
