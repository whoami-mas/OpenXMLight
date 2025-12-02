using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.SettingsW
{
    internal class StylesWord
    {
        internal Styles Styles { get; private set; }

        internal void GenerateStyles(MainDocumentPart mainPart)
        {
            StylesPart stylesPart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();

            Styles = stylesPart.Styles ??= new Styles();
        }

        internal string CreateGetEndnoteStyle()
        {
            Style? styleEndnote = Styles.OfType<Style>().FirstOrDefault(f => f.Elements<Name>().FirstOrDefault()?.Val == "endnote reference");

            if(styleEndnote == null)
            {
                int countStyle = Styles.Select(s => s.Elements<Style>()).Count();

                styleEndnote = new Style(
                    new Name() { Val = "endnote reference" },
                    new UIPriority() { Val = 99 },
                    new SemiHidden(),
                    new UnhideWhenUsed(),
                    new StyleRunProperties(
                        new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript }
                    )
                )
                {Type = StyleValues.Character, StyleId = $"a{countStyle}" };

                Styles.AppendChild(styleEndnote);
            }

            return styleEndnote.StyleId;
        }
    }
}
