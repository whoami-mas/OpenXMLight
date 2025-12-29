using OpenXMLight.Configurations.Elements.Interfaces;
using OpenXMLight.Configurations.Formatting;
using OpenXMLight.Tools.ToolsBase;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml;

namespace OpenXMLight.Configurations.Elements.TableElements.Formattings.FormattingsBase
{
    public abstract class WidthBase : ConvertBase
    {
        internal bool IsTypeFirst = false;

        public abstract string Width { get; set; }
        public abstract TypeWidthTable Type { get; set; }

        internal abstract void Initialize(OpenXml.OpenXmlElement? xml);
    }
}
