using OpenXMLight.Configurations.Elements.Interfaces;
using OpenXMLight.Configurations.Elements.TableElements.Formattings.FormattingsBase;
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
    public class Margin<T> : IMargin
        where T : MarginBase, new()
    {
        public string Width { get => _value.Width; set => _value.Width = value; }
        public TypeWidthTable Type => _value.Type;


        internal readonly T _value;
        internal Margin(OpenXml.OpenXmlElement? xml = null)
        {
            _value = new();
            _value.Initialize(xml);
        }
    }
}
