using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXML = DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlF = DocumentFormat.OpenXml;
using OpenXMLDrawing = DocumentFormat.OpenXml.Drawing;
using OpenXmlChart = DocumentFormat.OpenXml.Drawing.Charts;
using System.Globalization;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace OpenXMLight.Configurations.Elements.Charts
{
    public class ColumnChart : ChartBuilder
    {
        internal int[] axisID;

        internal override void GeneratedPlotArea()
        {
            OpenXmlChart.PlotArea plotAreaElement = new OpenXmlChart.PlotArea(
                new OpenXmlChart.Layout()
                );

            //BarChart
            OpenXmlChart.BarChart barChart = new OpenXmlChart.BarChart(
                new OpenXmlChart.BarDirection() { Val = OpenXmlChart.BarDirectionValues.Column },
                new OpenXmlChart.BarGrouping() { Val = OpenXmlChart.BarGroupingValues.Clustered },
                new OpenXmlChart.VaryColors() { Val = false }
                );

            for (int i = 0; i < this.Chart.Data.Count; i++)
            {
                //Index
                OpenXmlChart.BarChartSeries barSeries = new OpenXmlChart.BarChartSeries(
                    new OpenXmlChart.Index() { Val = Convert.ToUInt32(i) },
                    new OpenXmlChart.Order() { Val = Convert.ToUInt32(i) }
                    );

                //SeriesText
                barSeries.AppendChild(
                    new OpenXmlChart.SeriesText(
                        new OpenXmlChart.StringReference(
                            new OpenXmlChart.Formula() { Text = $"Лист1!${(char)('B' + i)}$1" },
                            new OpenXmlChart.StringCache(
                                new OpenXmlChart.PointCount() { Val = 1U },
                                new OpenXmlChart.StringPoint(
                                    new OpenXmlChart.NumericValue() { Text = chartData[i].Title }
                                    )
                                { Index = 0U }
                                )
                            )
                        )
                );

                //CategoryAxisDate
                OpenXmlChart.CategoryAxisData categoryAxisDate = new OpenXmlChart.CategoryAxisData();
                OpenXmlChart.StringReference strReference = new OpenXmlChart.StringReference(
                    new OpenXmlChart.Formula()
                    {
                        Text = chartData[i].Labels.Count() > 1 ? $"Лист1!$A$2:$A${chartData[i].Labels.Count() + 1}"
                                                                                        : "Лист1!$A$2"
                    }
                    );
                OpenXmlChart.StringCache strCache = new OpenXmlChart.StringCache(
                    new OpenXmlChart.PointCount() { Val = Convert.ToUInt32(chartData[i].Labels.Count()) }
                    );
                for (int j = 0; j < chartData[i].Data.Count(); j++)
                    strCache.AppendChild(
                        new OpenXmlChart.StringPoint(
                            new OpenXmlChart.NumericValue() { Text = chartData[i].Labels[j] }
                            )
                        { Index = Convert.ToUInt32(j) }
                        );
                strReference.AppendChild(strCache);
                categoryAxisDate.AppendChild(strReference);
                barSeries.AppendChild(categoryAxisDate);

                //Values
                OpenXmlChart.Values values = new OpenXmlChart.Values();
                OpenXmlChart.NumberReference numberReference = new OpenXmlChart.NumberReference(
                    new OpenXmlChart.Formula()
                    {
                        Text = chartData[i].Data.Count() > 1 ? $"Лист1!${(char)('B' + i)}$2:${(char)('B' + i)}${chartData[i].Data.Count() + 1}"
                                                                                      : $"Лист1!${(char)('B' + i)}$"
                    }
                    );
                OpenXmlChart.NumberingCache numCache = new OpenXmlChart.NumberingCache(
                    new OpenXmlChart.FormatCode() { Text = "General" },
                    new OpenXmlChart.PointCount() { Val = Convert.ToUInt32(chartData[i].Data.Count()) }
                    );
                for (int j = 0; j < chartData[i].Data.Count(); j++)
                    numCache.AppendChild(
                        new OpenXmlChart.NumericPoint(
                            new OpenXmlChart.NumericValue() { Text = chartData[i].Data[j].ToString(CultureInfo.InvariantCulture) }
                            )
                        { Index = Convert.ToUInt32(j) }
                        );
                numberReference.AppendChild(numCache);
                values.AppendChild(numberReference);

                barSeries.AppendChild(values);
                barChart.AppendChild(barSeries);
            }

            barChart.Append(
               new OpenXmlChart.AxisId() { Val = Convert.ToUInt32(axisID[0]) },
               new OpenXmlChart.AxisId() { Val = Convert.ToUInt32(axisID[1]) }
               );

            plotAreaElement.AppendChild(barChart);

            //CategoryAxis
            OpenXmlChart.CategoryAxis catAxis = new OpenXmlChart.CategoryAxis(
                new OpenXmlChart.AxisId() { Val = Convert.ToUInt32(axisID[0]) },
                new OpenXmlChart.Scaling(
                    new OpenXmlChart.Orientation() { Val = OpenXmlChart.OrientationValues.MinMax }
                ),
                new OpenXmlChart.Delete() { Val = false },
                new OpenXmlChart.AxisPosition() { Val = OpenXmlChart.AxisPositionValues.Bottom },
                new OpenXmlChart.NumberingFormat() { FormatCode = "General", SourceLinked = true },
                new OpenXmlChart.MajorTickMark() { Val = OpenXmlChart.TickMarkValues.None },
                new OpenXmlChart.MinorTickMark() { Val = OpenXmlChart.TickMarkValues.None },
                new OpenXmlChart.TickLabelPosition() { Val = OpenXmlChart.TickLabelPositionValues.NextTo },

                new OpenXmlChart.CrossingAxis() { Val = Convert.ToUInt32(axisID[1]) },
                new OpenXmlChart.Crosses() { Val = OpenXmlChart.CrossesValues.AutoZero },
                new OpenXmlChart.AutoLabeled() { Val = true },
                new OpenXmlChart.LabelAlignment() { Val = OpenXmlChart.LabelAlignmentValues.Center },
                new OpenXmlChart.LabelOffset() { Val = Convert.ToUInt16(100) },
                new OpenXmlChart.NoMultiLevelLabels() { Val = false }
            );
            plotAreaElement.AppendChild(catAxis);

            //ValueAxis
            OpenXmlChart.ValueAxis valAxis = new OpenXmlChart.ValueAxis(
                new OpenXmlChart.AxisId() { Val = Convert.ToUInt32(axisID[1]) },
                new OpenXmlChart.Scaling(
                    new OpenXmlChart.Orientation() { Val = OpenXmlChart.OrientationValues.MinMax }
                ),
                new OpenXmlChart.Delete() { Val = false },
                new OpenXmlChart.AxisPosition() { Val = OpenXmlChart.AxisPositionValues.Left },
                new OpenXmlChart.MajorGridlines(
                    new OpenXmlChart.ChartShapeProperties(
                        new OpenXMLDrawing.Outline(
                            new OpenXMLDrawing.SolidFill(
                                new OpenXMLDrawing.SchemeColor(
                                    new OpenXMLDrawing.LuminanceModulation() { Val = 15000 },
                                    new OpenXMLDrawing.LuminanceOffset() { Val = 85000 }
                                )
                                { Val = OpenXMLDrawing.SchemeColorValues.Text1 }
                            ),
                            new OpenXMLDrawing.Round()
                        )
                        {
                            Width = 9525,
                            CapType = OpenXMLDrawing.LineCapValues.Flat,
                            CompoundLineType = OpenXMLDrawing.CompoundLineValues.Single,
                            Alignment = OpenXMLDrawing.PenAlignmentValues.Center
                        },
                        new OpenXMLDrawing.EffectList()
                    )
                ),
                new OpenXmlChart.NumberingFormat() { FormatCode = "General", SourceLinked = true },
                new OpenXmlChart.MajorTickMark() { Val = OpenXmlChart.TickMarkValues.None },
                new OpenXmlChart.MinorTickMark() { Val = OpenXmlChart.TickMarkValues.None },
                new OpenXmlChart.TickLabelPosition() { Val = OpenXmlChart.TickLabelPositionValues.NextTo },
                new OpenXmlChart.CrossingAxis() { Val = Convert.ToUInt32(axisID[0]) },
                new OpenXmlChart.Crosses() { Val = OpenXmlChart.CrossesValues.AutoZero },
                new OpenXmlChart.CrossBetween() { Val = OpenXmlChart.CrossBetweenValues.Between }
            );
            plotAreaElement.AppendChild(valAxis);

            this.Chart.ChartXml.AppendChild(plotAreaElement);
        }

        internal override void GeneratedLegend()
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
