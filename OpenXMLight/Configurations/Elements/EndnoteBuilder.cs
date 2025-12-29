using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements
{
    public class EndnoteBuilder
    {
        EndnoteTest endnote;


        internal EndnoteBuilder(OpenXml.Endnote endnote) => this.endnote = new(endnote);

        public EndnoteBuilder AddParagraph(Action<ParagraphBuilder>? configuration = null)
        {
            OpenXml.Paragraph p = new OpenXml.Paragraph();
            endnote.ElementXml.AppendChild(p);

            var paragraph = new ParagraphBuilder(p);
            configuration?.Invoke(paragraph);

            return this;
        }


        public static implicit operator EndnoteTest(EndnoteBuilder build) => build.endnote;
    }
}
