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
        private static Context? _instance = null;


        internal static Context GetInstance(OpenXmlPackaging.MainDocumentPart? mainDoc = null)
        {
            if (_instance == null)
                _instance = new(mainDoc);

            return _instance;
        }


        public Styles Styles { get; init; }
        public Endnotes Endnotes { get; init; }


        protected Context(OpenXmlPackaging.MainDocumentPart? mainDoc = null)
        {
            if(mainDoc != null)
            {
                Styles = new(mainDoc.StyleDefinitionsPart ?? mainDoc.AddNewPart<OpenXmlPackaging.StyleDefinitionsPart>());
                Endnotes = new(mainDoc.EndnotesPart ?? mainDoc.AddNewPart<OpenXmlPackaging.EndnotesPart>());
            }
        }
    }
}
