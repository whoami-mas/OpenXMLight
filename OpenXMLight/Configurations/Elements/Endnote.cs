using OpenXMLight.Configurations.Elements.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements
{
    public class Endnote : Element<OpenXml.Endnote, OpenXml.EndnoteProperties>
    {
        internal override OpenXml.Endnote ElementXml { get; set; }
        internal override OpenXml.EndnoteProperties ElementProperties
        {
            get
            {
                if (_elementProperties == null)
                {
                    _elementProperties = ElementXml.Elements<OpenXml.EndnoteProperties>().FirstOrDefault();

                    if (_elementProperties == null)
                        _elementProperties = ElementXml.PrependChild(new OpenXml.EndnoteProperties());
                }

                return _elementProperties;
            }
        }


        internal Endnote(OpenXml.Run r)
        {
            long id_endnote = r.Elements<OpenXml.EndnoteReference>().First().Id;


        }
        internal Endnote(OpenXml.Endnote e) => ElementXml = e;


        #region Private properties
        private OpenXml.EndnoteProperties? _elementProperties;
        private ElementCollection<Paragraph> paragraphs;
        #endregion

        public ElementCollection<Paragraph> Paragraphs
        {
            get
            {
                if(paragraphs == null)
                    paragraphs = new(ElementXml.Elements<OpenXml.Paragraph>().Select(s => new Paragraph(s))) { Parent = ElementXml };

                return paragraphs;
            }
        }
    }
}
