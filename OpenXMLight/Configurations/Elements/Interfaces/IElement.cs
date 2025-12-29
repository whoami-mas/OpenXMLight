using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Configurations.Elements.Interfaces
{
    public interface IElement
    {
        OpenXmlElement XmlElement { get; }
    }
}
