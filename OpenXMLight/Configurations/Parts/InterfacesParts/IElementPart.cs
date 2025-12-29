using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;

namespace OpenXMLight.Configurations.Parts.InterfacesParts
{
    public interface IElementPart<T> where T : OpenXmlPackaging.OpenXmlPart 
    {
        T PartXml { get; set;  }
    }
}
