using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXMLDrawing = DocumentFormat.OpenXml.Drawing;
using OpenXmlChart = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLight.Configurations.Elements.Charts
{
    public abstract class ChartBuilder
    {
        public Chart Chart { get; private set; }

        public ChartBuilder SetTitle(string title)
        {
            Chart.Title = title;

            return this;
        }
        public ChartBuilder SetData(List<ChartData> data)
        {
            Chart.Data = data;

            return this;
        }

        public ChartBuilder()
        {
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
                            new OpenXMLDrawing.Run(
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
        internal abstract void GeneratedLegend();
    }
}
