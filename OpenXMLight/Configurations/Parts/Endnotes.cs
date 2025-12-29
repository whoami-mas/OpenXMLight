using OpenXMLight.Configurations.Parts.InterfacesParts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXMLight.Configurations.WordContext;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXml = DocumentFormat.OpenXml.Wordprocessing;
using Formatte = DocumentFormat.OpenXml;

namespace OpenXMLight.Configurations.Parts
{
    internal class Endnotes : IElementPart<OpenXmlPackaging.EndnotesPart>
    {
        public OpenXmlPackaging.EndnotesPart PartXml { get; set; }

        internal Endnotes(OpenXmlPackaging.EndnotesPart endnotesPart) => PartXml = endnotesPart;

        internal long GetMaxId()
        {
            long id = 1;

            if(PartXml.Endnotes != null && PartXml.Endnotes.Count() > 0)
                id = PartXml.Endnotes.OfType<OpenXml.Endnote>().Select(w => w.Id).DefaultIfEmpty(1).Max();

            return id;
        }
        public OpenXml.Endnote AddEndnote(string content)
        {
            CheckedExists();

            OpenXml.Endnote endnote = new OpenXml.Endnote(
                new OpenXml.Paragraph(
                    new OpenXml.ParagraphProperties(
                        new OpenXml.ParagraphStyleId() { Val = Context.GetInstance().Styles.CreateGetEndnoteText() }
                        ),
                    new OpenXml.Run(
                        new OpenXml.RunProperties(
                            new OpenXml.RunStyle() { Val = Context.GetInstance().Styles.CreateGetEndnoteRef() }
                            ),
                        new OpenXml.EndnoteReferenceMark()
                        ),
                    new OpenXml.Run(
                        new OpenXml.Text(" ") { Space = Formatte.SpaceProcessingModeValues.Preserve }
                        ),
                    new OpenXml.Run(
                        new OpenXml.Text(content)
                        )
                    )
                ) { Id = GetMaxId() + 1 };

            PartXml.Endnotes.AppendChild(endnote);
            
            return endnote;
        }

        private void CheckedExists()
        {
            PartXml.Endnotes ??= new OpenXml.Endnotes();

            PartXml.Endnotes.AppendChild(
                new OpenXml.Endnote(
                     new OpenXml.Paragraph(
                        new OpenXml.ParagraphProperties(
                            new OpenXml.SpacingBetweenLines() { After = "0", Line = "240", LineRule = OpenXml.LineSpacingRuleValues.Auto}
                            ),
                        new OpenXml.Run(
                            new OpenXml.SeparatorMark()
                            )
                        )
                     )
                { Type = OpenXml.FootnoteEndnoteValues.Separator, Id = -1 }
                );
            PartXml.Endnotes.AppendChild(
                new OpenXml.Endnote(
                     new OpenXml.Paragraph(
                        new OpenXml.ParagraphProperties(
                            new OpenXml.SpacingBetweenLines() { After = "0", Line = "240", LineRule = OpenXml.LineSpacingRuleValues.Auto }
                            ),
                        new OpenXml.Run(
                            new OpenXml.ContinuationSeparatorMark()
                            )
                        )
                     )
                { Type = OpenXml.FootnoteEndnoteValues.ContinuationSeparator, Id = 0 }
                );
        }
    }
}
