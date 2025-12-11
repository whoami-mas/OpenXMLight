using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLight.Configurations;
using OpenXMLight.Configurations.Formatting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.SettingsW
{
    public class SettingsPageWord
    {
        private OrientationPage orientation;
        
        private int marginTop;
        private int marginLeft;
        private int marginRight;
        private int marginBottom;

        private int width;
        private int height;

        //pgSz
        public int WidthPage { get => width / Configuration.TwipsInPixels; 
            set {
                width = value * Configuration.TwipsInPixels;

                SectionPropXml.GetFirstChild<PageSize>().Width = Convert.ToUInt32(width);
            } }
        public int HeightPage { get => height / Configuration.TwipsInPixels;
            set {
                height = value * Configuration.TwipsInPixels;

                SectionPropXml.GetFirstChild<PageSize>().Height = Convert.ToUInt32(height);
            } }

        //Orient
        public OrientationPage Orientation 
        { 
            get => orientation;
            set { 
                orientation = value;
                SectionPropXml.GetFirstChild<PageSize>().Orient = value.Value;

                if (value.Value == PageOrientationValues.Landscape)
                {
                    MarginTop = 113;
                    MarginRight = 75;
                    MarginBottom = 56;
                    MarginLeft = 75;

                    WidthPage = 1122;
                    HeightPage = 793;
                }
                else if(value.Value == PageOrientationValues.Portrait)
                {
                    MarginTop = 75;
                    MarginRight = 56;
                    MarginBottom = 75;
                    MarginLeft = 113;

                    WidthPage = 793;
                    HeightPage = 1122;
                }
            } 
        }

        //pgMar
        public int MarginTop { get => marginTop / Configuration.TwipsInPixels;
            set { 
                marginTop = value * Configuration.TwipsInPixels;
                SectionPropXml.GetFirstChild<PageMargin>().Top = marginTop;
            } 
        }
        public int MarginLeft { get => marginLeft / Configuration.TwipsInPixels;
            set {
                marginLeft = value * Configuration.TwipsInPixels;
                SectionPropXml.GetFirstChild<PageMargin>().Left = Convert.ToUInt32(marginLeft);
            } }
        public int MarginRight { get => marginRight / Configuration.TwipsInPixels;
            set {
                marginRight = value * Configuration.TwipsInPixels;
                SectionPropXml.GetFirstChild<PageMargin>().Right = Convert.ToUInt32(marginRight);
            } }
        public int MarginBottom { get => marginBottom / Configuration.TwipsInPixels;
            set {
                marginBottom = value * Configuration.TwipsInPixels;
                SectionPropXml.GetFirstChild<PageMargin>().Bottom = marginBottom;
            } }
        public int MarginHeader { get; set; }
        public int MarginFooter { get; set; }
        public int MarginGutter { get; set; }

        //cols
        public string? Space { get; set; } = null;

        //docGrid
        public int LinePitch { get; set; }

        internal SectionProperties? SectionPropXml { get; set; }

        internal void GenerateDocumentSettings(Document document)
        {
            SectionPropXml = document.Body.Elements<SectionProperties>().FirstOrDefault();

            SectionPropXml ??= document.Body.AppendChild(
                            new SectionProperties(
                                        new PageSize() { Width = (UInt32Value)11906U, Height = (UInt32Value)16838U, Orient = Orientation.Value },
                                        new PageMargin() { Top = 1134, Right = (UInt32Value)850U, Bottom = 1134, Left = (UInt32Value)1701U, Header = (UInt32Value)708U, Footer = (UInt32Value)708U, Gutter = (UInt32Value)0U },
                                        new Columns() { Space = "708" },
                                        new DocGrid() { LinePitch = 360 }
                                        )
                            );

            PageSize pgSize = SectionPropXml.Elements<PageSize>().First();
            WidthPage = (int)pgSize.Width.Value;
            HeightPage = (int)pgSize.Height.Value;

            PageMargin pgMargin = SectionPropXml.Elements<PageMargin>().First();
            MarginTop = pgMargin.Top;
            MarginLeft = (int)pgMargin.Left.Value;
            MarginRight = (int)pgMargin.Right.Value;
            MarginBottom = (int)pgMargin.Bottom.Value;
            MarginHeader = (int)pgMargin.Header.Value;
            MarginFooter = (int)pgMargin.Footer.Value;
            MarginGutter = (int)pgMargin.Gutter.Value;

            Columns col = SectionPropXml.Elements<Columns>().First();
            Space = col.Space;

            DocGrid docGrid = SectionPropXml.Elements<DocGrid>().First();
            LinePitch = docGrid.LinePitch;
        }
    }
}
