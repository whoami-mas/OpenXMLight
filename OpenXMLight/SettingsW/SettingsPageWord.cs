using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.SettingsW
{
    internal class SettingsPageWord
    {
        //pgSz
        internal UInt32Value? WidthPage { get; set; } = null;
        internal UInt32Value? HeightPage { get; set; } = null;
        
        //pgMar
        internal int? MarginTop { get; set; } = null;
        internal UInt32Value? MarginLeft { get; set; } = null;
        internal UInt32Value? MarginRight { get; set; } = null;
        internal int? MarginBottom { get; set; } = null;
        internal UInt32Value? MarginHeader { get; set; } = null;
        internal UInt32Value? MarginFooter { get; set; } = null;
        internal UInt32Value? MarginGutter { get; set; } = null;

        //cols
        internal string? Space { get; set; } = null;

        //docGrid
        internal int? LinePitch { get; set; } = null;

        internal void GenerateDocumentSettings(Document document)
        {
            SectionProperties? secProp = document.Body.Elements<SectionProperties>().FirstOrDefault();

            secProp ??= document.Body.AppendChild(
                            new SectionProperties(
                                        new PageSize() { Width = (UInt32Value)11906U, Height = (UInt32Value)16838U},
                                        new PageMargin() { Top = 1134, Right = (UInt32Value)850U, Bottom = 1134, Left = (UInt32Value)1701U, Header = (UInt32Value)708U, Footer = (UInt32Value)708U, Gutter = (UInt32Value)0U },
                                        new Columns() { Space = "708" },
                                        new DocGrid() { LinePitch = 360 }
                                        )
                            );

            PageSize pgSize = secProp.Elements<PageSize>().First();
            WidthPage = pgSize.Width;
            HeightPage = pgSize.Height;

            PageMargin pgMargin = secProp.Elements<PageMargin>().First();
            MarginTop = pgMargin.Top;
            MarginLeft = pgMargin.Left;
            MarginRight = pgMargin.Right;
            MarginBottom = pgMargin.Bottom;
            MarginFooter = pgMargin.Footer;
            MarginGutter = pgMargin.Gutter;

            Columns col = secProp.Elements<Columns>().First();
            Space = col.Space;

            DocGrid docGrid = secProp.Elements<DocGrid>().First();
            LinePitch = docGrid.LinePitch;
        }
    }
}
