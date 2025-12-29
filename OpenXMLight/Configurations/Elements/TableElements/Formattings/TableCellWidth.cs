using OpenXMLight.Configurations.Elements.TableElements.Formattings.FormattingsBase;
using OpenXMLight.Configurations.Formatting;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml;

namespace OpenXMLight.Configurations.Elements.TableElements.Formattings
{
    public class TableCellWidth<T> where T : WidthBase, new()
    {
        public string Width
        {
            get => _value.Width;
            set => _value.Width = value;
        }
        public TypeWidthTable Type
        {
            get => _value.Type;
            set => _value.Type = value;
        }


        internal readonly T _value;

        internal TableCellWidth(OpenXml.OpenXmlElement? xml = null)
        {
            _value = new T();
            _value.Initialize(xml);
        }

    }
}
