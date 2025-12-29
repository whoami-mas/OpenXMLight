using DocumentFormat.OpenXml;
using OpenXMLight.Configurations.Elements.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Configurations.Elements
{
    public abstract class Element<TElement, TProperties> : IElement 
        where TElement : OpenXmlElement
        where TProperties : OpenXmlElement
    {
        internal abstract TElement ElementXml { get; set; }
        internal abstract TProperties ElementProperties { get; }

        OpenXmlElement IElement.XmlElement => ElementXml;
    }
}
