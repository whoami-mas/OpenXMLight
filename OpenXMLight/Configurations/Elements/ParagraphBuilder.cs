using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlFormatte = DocumentFormat.OpenXml;

using OpenXMLight.Configurations.Formatting;
using OpenXMLight.config;

namespace OpenXMLight.Configurations.Elements
{
    public class ParagraphBuilder
    {
        Paragraph _paragraph;



        public ParagraphBuilder() => _paragraph = new(new OpenXml.Paragraph());
        internal ParagraphBuilder(OpenXml.Paragraph p) => _paragraph = new(p);



        public ParagraphBuilder SetAlignment(HorizontalAlignments alignment)
        {
            _paragraph.ElementProperties.Justification = new OpenXml.Justification() { Val = alignment.Value };

            return this;
        }
        public ParagraphBuilder SetRun(params RunBuilder[] runs)
        {
            foreach(RunBuilder run in runs)
            {
                Run r = run;
                _paragraph.ElementXml.Append(r.ElementXml);
            }

            return this;
        }
        public ParagraphBuilder SetSpacingBetweenLines(SpacingBetweenLines spacing)
        {
            _paragraph.ElementProperties.SpacingBetweenLines = new OpenXml.SpacingBetweenLines() {
                After = spacing.After.ToString(),
                Before = spacing.Before.ToString(),
                Line = spacing.Line.ToString()
            };

            return this;
        }



        public static implicit operator Paragraph(ParagraphBuilder builder)
        {
            return builder._paragraph;
        }
    }
}
