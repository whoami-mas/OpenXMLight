using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXMLDrawing = DocumentFormat.OpenXml.Drawing;
using OpenXmlChart = DocumentFormat.OpenXml.Drawing.Charts;
using OpenXMLight.Spreadsheet.Formatting;

namespace OpenXMLight.Configurations.Elements.Charts
{
    internal enum Axes
    {
        Left, Top, Right, Bottom
    }
    public abstract class ChartBuilder
    {
        public Chart Chart { get; private set; }
        internal Dictionary<int, OpenXMLDrawing.SchemeColorValues> StyleLine { get; private set; }
        internal Dictionary<Axes, int> AxesID { get; init; }
        internal bool IsRightAxis { get; set; } = false;
        internal TypeValue TypeFormatAxis { get; set; }

        public ChartBuilder SetTitle(string title)
        {
            Chart.Title = title;

            return this;
        }
        public virtual ChartBuilder SetData(List<ChartData> data)
        {
            Chart.Data = data;

            return this;
        }
        public ChartBuilder SetSize(int width, int height)
        {
            this.Chart.Width = width;
            this.Chart.Height = height;

            return this;
        }
        public ChartBuilder SetIsRightAxis(bool isRightAxis, TypeValue TypeFormat)
        {
            IsRightAxis = isRightAxis;
            TypeFormatAxis = TypeFormat;

            return this;
        }


        public ChartBuilder()
        {
            StyleLine = new()
            {
                {0, OpenXMLDrawing.SchemeColorValues.Accent1 },
                {1, OpenXMLDrawing.SchemeColorValues.Accent2 },
                {2, OpenXMLDrawing.SchemeColorValues.Accent3 },
                {3, OpenXMLDrawing.SchemeColorValues.Accent4 },
                {4, OpenXMLDrawing.SchemeColorValues.Accent5 },
                {5, OpenXMLDrawing.SchemeColorValues.Accent6 },
            };

            Chart = new Chart();
        }


        internal virtual void GeneratedTitle()
        {
            OpenXmlChart.Title titleElement = new OpenXmlChart.Title();

            titleElement.AppendChild(
                new OpenXmlChart.ChartText(
                    new OpenXmlChart.RichText(
                        new OpenXMLDrawing.BodyProperties()
                        {
                            Anchor = OpenXMLDrawing.TextAnchoringTypeValues.Center,
                            AnchorCenter = true,
                            Rotation = 0,
                            UseParagraphSpacing = true,
                            VerticalOverflow = OpenXMLDrawing.TextVerticalOverflowValues.Ellipsis,
                            Vertical = OpenXMLDrawing.TextVerticalValues.Horizontal,
                            Wrap = OpenXMLDrawing.TextWrappingValues.Square
                        },
                        new OpenXMLDrawing.Paragraph(
                            new OpenXMLDrawing.ParagraphProperties(
                                new OpenXMLDrawing.DefaultRunProperties(
                                    new OpenXMLDrawing.SolidFill(
                                        new OpenXMLDrawing.SchemeColor(
                                            new OpenXMLDrawing.LuminanceModulation() { Val = 65000 },
                                            new OpenXMLDrawing.LuminanceOffset() { Val = 35000 }
                                        ) { Val = OpenXMLDrawing.SchemeColorValues.Text1}
                                    )
                                )
                                {
                                    FontSize = 1400,
                                    Bold = false,
                                    Italic = false,
                                    Underline = OpenXMLDrawing.TextUnderlineValues.None,
                                    Strike = OpenXMLDrawing.TextStrikeValues.NoStrike,
                                    Kerning = 1200,
                                    Spacing = 0,
                                    Baseline = 0
                                }
                            ),
                            new OpenXMLDrawing.Run(
                                new OpenXMLDrawing.RunProperties() { FontSize = 1500 },
                                new OpenXMLDrawing.Text(this.Chart.Title)
                            )
                        )
                    )
                )
            );

            titleElement.Append(
                new OpenXmlChart.Overlay() { Val = false },
                new OpenXmlChart.TextProperties(
                    new OpenXMLDrawing.BodyProperties()
                    {
                        Rotation = 0,
                        UseParagraphSpacing = true,
                        VerticalOverflow = OpenXMLDrawing.TextVerticalOverflowValues.Ellipsis,
                        Vertical = OpenXMLDrawing.TextVerticalValues.Horizontal,
                        Wrap = OpenXMLDrawing.TextWrappingValues.Square,
                        Anchor = OpenXMLDrawing.TextAnchoringTypeValues.Center,
                        AnchorCenter = true
                    },
                    new OpenXMLDrawing.Paragraph(
                        )
                )
            );

            this.Chart.ChartXml.Append(titleElement);
        }
        internal virtual void GeneratedAutoTitleDeleted() => this.Chart.ChartXml.AppendChild(new OpenXmlChart.AutoTitleDeleted() { Val = false});
        internal abstract void GeneratedPlotArea();
        internal virtual void GeneratedPlotVisibleOnly() => this.Chart.ChartXml.AppendChild(new OpenXmlChart.PlotVisibleOnly() { Val = true});
        internal virtual void GeneratedLegend()
        {
            OpenXmlChart.Legend legendElement = new OpenXmlChart.Legend(
            new OpenXmlChart.LegendPosition() { Val = OpenXmlChart.LegendPositionValues.Bottom },
            new OpenXmlChart.Overlay() { Val = false },
                new OpenXmlChart.TextProperties(
                    new OpenXMLDrawing.BodyProperties()
                    {
                        Rotation = 0,
                        UseParagraphSpacing = true,
                        VerticalOverflow = OpenXMLDrawing.TextVerticalOverflowValues.Ellipsis,
                        Vertical = OpenXMLDrawing.TextVerticalValues.Horizontal,
                        Wrap = OpenXMLDrawing.TextWrappingValues.Square,
                        Anchor = OpenXMLDrawing.TextAnchoringTypeValues.Center,
                        AnchorCenter = true
                    },
                    new OpenXMLDrawing.Paragraph(
                        new OpenXMLDrawing.ParagraphProperties(
                            new OpenXMLDrawing.DefaultRunProperties(
                                new OpenXMLDrawing.SolidFill(
                                    new OpenXMLDrawing.SchemeColor(
                                        new OpenXMLDrawing.LuminanceModulation() { Val = 65000 },
                                        new OpenXMLDrawing.LuminanceOffset() { Val = 35000 }
                                        )
                                    { Val = OpenXMLDrawing.SchemeColorValues.Text1 }
                                    )
                                )
                            {
                                FontSize = 900,
                                Bold = false,
                                Italic = false,
                                Underline = OpenXMLDrawing.TextUnderlineValues.None,
                                Strike = OpenXMLDrawing.TextStrikeValues.NoStrike,
                                Kerning = 1200,
                                Baseline = 0
                            }
                            )
                        )
                )
            );

            this.Chart.ChartXml.AppendChild(legendElement);
        }
    }
}
