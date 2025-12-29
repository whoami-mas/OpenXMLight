using OpenXMLight.Configurations.Elements.Interfaces;
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

namespace OpenXMLight.Configurations.Elements.TableElements.Formattings
{
    public class TableCellMargin<T> where T : IMarginContainer, new()
    {
        public IMargin Top => _value.Top;
        public IMargin Bottom => _value.Bottom;
        public IMargin Left => _value.Left;
        public IMargin Right => _value.Right;


        internal readonly T _value;
        internal TableCellMargin(OpenXml.OpenXmlElement? xml = null)
        {
            _value = new();
            _value.Initialize(xml);
        }
    }
}
