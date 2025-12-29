using OpenXMLight.Configurations.Elements.Interfaces;
using OpenXMLight.Configurations.Elements.TableElements.Formattings.MarginComponents;
using OpenXMLight.Configurations.Formatting;
using OpenXMLight.Tools.ToolsBase;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlElement = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements.TableElements.Formattings.FormattingsBase
{
    public abstract class MarginBase : ConvertBase
    {
        public abstract string Width { get; set; }
        public abstract TypeWidthTable Type { get; }

        internal abstract void Initialize(OpenXml.OpenXmlElement? xml);
    }
}
