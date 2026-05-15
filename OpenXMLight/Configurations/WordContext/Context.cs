using OpenXMLight.Configurations.Parts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;

namespace OpenXMLight.Configurations.WordContext
{
    internal class Context : IContext
    {
        public Styles Styles { get; init; }
        public Endnotes? Endnotes { get; init; }
        

        internal Context(OpenXmlPackaging.MainDocumentPart? mainDoc = null)
        {
            if(mainDoc != null)
            {
                Styles = new(mainDoc.StyleDefinitionsPart ?? mainDoc.AddNewPart<OpenXmlPackaging.StyleDefinitionsPart>());
                Endnotes = new(mainDoc.EndnotesPart ?? mainDoc.AddNewPart<OpenXmlPackaging.EndnotesPart>(), this);
            }
        }
    }
}
