using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
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
        public int WidthPage { get => width; 
            set {
                width = value;

                SectionPropXml.GetFirstChild<PageSize>().Width = Convert.ToUInt32(value);
            } }
        public int HeightPage { get => height;
            set {
                height = value;

                SectionPropXml.GetFirstChild<PageSize>().Height = Convert.ToUInt32(value);
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
                    MarginTop = 1701;
                    MarginRight = 1134;
                    MarginBottom = 850;
                    MarginLeft = 1134;

                    WidthPage = 16838;
                    HeightPage = 11907;
                }
                else if(value.Value == PageOrientationValues.Portrait)
                {
                    MarginTop = 1134;
                    MarginRight = 850;
                    MarginBottom = 1134;
                    MarginLeft = 1701;

                    WidthPage = 11907;
                    HeightPage = 16838;
                }
            } 
        }

        //pgMar
        public int MarginTop { get => marginTop;
            set { 
                marginTop = value;
                SectionPropXml.GetFirstChild<PageMargin>().Top = value;
            } 
        }
        public int MarginLeft { get => marginLeft;
            set {
                marginLeft = value;
                SectionPropXml.GetFirstChild<PageMargin>().Left = Convert.ToUInt32(value);
            } }
        public int MarginRight { get => marginRight;
            set {
                marginRight = value;
                SectionPropXml.GetFirstChild<PageMargin>().Right = Convert.ToUInt32(value);
            } }
        public int MarginBottom { get => marginBottom;
            set {
                marginBottom = value;
                SectionPropXml.GetFirstChild<PageMargin>().Bottom = value;
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
