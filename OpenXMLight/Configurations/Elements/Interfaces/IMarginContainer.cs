using OpenXMLight.Configurations.Elements.TableElements.Formattings.FormattingsBase;
using OpenXMLight.Configurations.Elements.TableElements.Formattings.MarginComponents;
using OpenXMLight.Configurations.Elements.TableElements.Formattings.MarginComponents.Margins;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlElement = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements.Interfaces
{
    public interface IMarginContainer
    {
        public IMargin Top { get; }
        public IMargin Bottom { get; }
        public IMargin Left { get; }
        public IMargin Right { get; }

        void Initialize(OpenXml.OpenXmlElement? xml);
    }
}
