using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml;

namespace OpenXMLight.Configurations.Elements.Interfaces
{
    internal interface IElementBase<T> where T : OpenXml.OpenXmlElement
    {
        public T ElementXml { get; set; }
    }
}
