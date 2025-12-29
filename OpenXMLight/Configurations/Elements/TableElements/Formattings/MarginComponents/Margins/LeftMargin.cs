using OpenXMLight.Configurations.Elements.Interfaces;
using OpenXMLight.Configurations.Elements.TableElements.Formattings.FormattingsBase;
using OpenXMLight.Configurations.Elements.TableElements.Formattings.MarginComponents;
using OpenXMLight.Configurations.Formatting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlElement = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements.TableElements.Formattings.MarginComponents.Margins
{
    public class LeftMargin : MarginBase, IElementBase<OpenXmlElement.LeftMargin>
    {
        private string width;

        public OpenXmlElement.LeftMargin ElementXml { get; set; }
        public override string Width
        {
            get => width;
            set
            {
                width = value;

                ElementXml.Width = GetWidthOfType<string>(width, Type);
            }
        }
        public override TypeWidthTable Type { get; } = TypeWidthTable.Cm;


        internal override void Initialize(OpenXml.OpenXmlElement? xml)
        {
            ElementXml = xml as OpenXmlElement.LeftMargin ?? new() { Type = Type.Value };

            width = ConvertWidth(ElementXml.Width, Type);
        }


        public static implicit operator OpenXmlElement.LeftMargin(LeftMargin build) => build.ElementXml;
    }
}
