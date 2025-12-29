using OpenXMLight.Configurations.Elements.Interfaces;
using OpenXMLight.Configurations.Parts.InterfacesParts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXml = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Parts
{
    internal class Styles : IElementPart<OpenXmlPackaging.StyleDefinitionsPart>
    {
        public OpenXmlPackaging.StyleDefinitionsPart PartXml { get; set; }
        private int CountStyles => PartXml.Styles.ChildElements.Count();

        internal Styles(OpenXmlPackaging.StyleDefinitionsPart stylesPart) => PartXml = stylesPart;


        private void CheckedExists() => PartXml.Styles ??= new OpenXml.Styles();

        internal string CreateGetEndnoteRef()
        {
            CheckedExists();

            OpenXml.Style? style = PartXml.Styles?.OfType<OpenXml.Style>().FirstOrDefault(f=> string.Equals(f.Elements<OpenXml.Name>().FirstOrDefault()?.Val, "endnote reference"));
            
            if(style == null)
            {
                style = new OpenXml.Style(
                    new OpenXml.Name() { Val = "endnote reference" },
                    new OpenXml.UIPriority() { Val = 99 },
                    new OpenXml.SemiHidden(),
                    new OpenXml.UnhideWhenUsed(),
                    new OpenXml.StyleRunProperties(
                        new OpenXml.VerticalTextAlignment() { Val = OpenXml.VerticalPositionValues.Superscript }
                        )
                    )
                { Type = OpenXml.StyleValues.Character, StyleId = $"a{CountStyles}" };

                PartXml.Styles?.AppendChild(style);
            }

            return style.StyleId;
        }
        internal string CreateGetEndnoteText()
        {
            CheckedExists();

            OpenXml.Style style = PartXml.Styles?.OfType<OpenXml.Style>().FirstOrDefault(f => string.Equals(f.Elements<OpenXml.Name>().FirstOrDefault()?.Val, "endnote text"));

            if (style == null)
            {
                style = new OpenXml.Style(
                    new OpenXml.Name() { Val = "endnote text" },
                    new OpenXml.UIPriority() { Val = 99 },
                    new OpenXml.SemiHidden(),
                    new OpenXml.UnhideWhenUsed(),
                    new OpenXml.StyleParagraphProperties(
                        new OpenXml.SpacingBetweenLines() { After = "0", Line = "240", LineRule = OpenXml.LineSpacingRuleValues.Auto}
                        ),
                    new OpenXml.StyleRunProperties(
                        new OpenXml.FontSize() { Val = "20"},
                        new OpenXml.FontSizeComplexScript() { Val = "20"}
                        )
                    )
                { Type = OpenXml.StyleValues.Paragraph, StyleId = $"a{CountStyles}" };

                PartXml.Styles?.AppendChild(style);
            }

            return style.StyleId;
        }
    }
}
