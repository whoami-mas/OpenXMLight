using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXmlF = DocumentFormat.OpenXml;
using OpenXMLDrawing = DocumentFormat.OpenXml.Drawing;
using OpenXmlChart = DocumentFormat.OpenXml.Drawing.Charts;
using System.Globalization;

namespace OpenXMLight.Configurations.Elements.Charts
{
    public class ColumnChart : ChartBuilder
    {
        internal int[] axisID;
        public ColumnChart()
        {
            axisID = new int[2] { Random.Shared.Next(100000000, 999999999), Random.Shared.Next(100000000, 999999999) };
        }

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
                                    new OpenXmlChart.NumericValue() { Text = this.Chart.Data[i].Title }
                                    )
                                { Index = 0U }
                                )
                            )
                        )
                );

                //ShapeProperties
                barSeries.AppendChild(
                    new OpenXmlChart.ChartShapeProperties(
                            new OpenXMLDrawing.SolidFill(
                                new OpenXMLDrawing.SchemeColor() { Val = this.StyleLine[i] }
                                ),
                            new OpenXMLDrawing.Outline(
                                new OpenXMLDrawing.NoFill()
                                ),
                            new OpenXMLDrawing.EffectList()
                        )
                    );

                barSeries.AppendChild(
                    new OpenXmlChart.InvertIfNegative() { Val = false}
                    );

                #region CategoryAxisDate
                OpenXmlChart.CategoryAxisData categoryAxisDate = new OpenXmlChart.CategoryAxisData();
                OpenXmlChart.StringReference strReference = new OpenXmlChart.StringReference(
                    new OpenXmlChart.Formula()
                    {
                        Text = this.Chart.Data[i].Labels.Length > 1 ? $"Лист1!$A$2:$A${this.Chart.Data[i].Labels.Length}"
                                                                     : "Лист1!$A$2"
                    }
                    );
                OpenXmlChart.StringCache strCache = new OpenXmlChart.StringCache(
                    new OpenXmlChart.PointCount() { Val = Convert.ToUInt32(this.Chart.Data[i].Labels.Length) }
                    );
                for (int j = 0; j < this.Chart.Data[i].Data.Length; j++)
                    strCache.AppendChild(
                        new OpenXmlChart.StringPoint(
                            new OpenXmlChart.NumericValue() { Text = this.Chart.Data[i].Labels[j] }
                            )
                        { Index = Convert.ToUInt32(j) }
                        );
                strReference.AppendChild(strCache);
                categoryAxisDate.AppendChild(strReference);
                barSeries.AppendChild(categoryAxisDate);
                #endregion

                #region Values
                OpenXmlChart.Values values = new OpenXmlChart.Values();
                OpenXmlChart.NumberReference numberReference = new OpenXmlChart.NumberReference(
                    new OpenXmlChart.Formula()
                    {
                        Text = this.Chart.Data[i].Data.Length > 1 ? $"Лист1!${(char)('B' + i)}$2:${(char)('B' + i)}${this.Chart.Data[i].Data.Length}"
                                                                   : $"Лист1!${(char)('B' + i)}$"
                    }
                    );
                OpenXmlChart.NumberingCache numCache = new OpenXmlChart.NumberingCache(
                    new OpenXmlChart.FormatCode() { Text = "General" },
                    new OpenXmlChart.PointCount() { Val = Convert.ToUInt32(this.Chart.Data[i].Data.Length) }
                    );
                for (int j = 0; j < this.Chart.Data[i].Data.Length; j++)
                    numCache.AppendChild(
                        new OpenXmlChart.NumericPoint(
                            new OpenXmlChart.NumericValue() { Text = this.Chart.Data[i].Data[j].ToString(CultureInfo.InvariantCulture) }
                            )
                        { Index = Convert.ToUInt32(j) }
                        );
                numberReference.AppendChild(numCache);
                values.AppendChild(numberReference);

                barSeries.AppendChild(values);
                #endregion

                barChart.AppendChild(barSeries);
            }

            barChart.Append(
                new OpenXmlChart.GapWidth() { Val = (OpenXmlF.UInt16Value)219U },
                new OpenXmlChart.Overlap() { Val = -27 }
                );

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
                new OpenXmlChart.ChartShapeProperties(
                    new OpenXMLDrawing.NoFill(),
                    new OpenXMLDrawing.Outline(
                        new OpenXMLDrawing.SolidFill(
                            new OpenXMLDrawing.SchemeColor(
                                new OpenXMLDrawing.LuminanceModulation() { Val = 15000 },
                                new OpenXMLDrawing.LuminanceOffset() { Val = 85000 }
                            )
                            { Val = OpenXMLDrawing.SchemeColorValues.Text1 }
                        )
                    )
                    {
                        Width = 9525,
                        CapType = OpenXMLDrawing.LineCapValues.Flat,
                        CompoundLineType = OpenXMLDrawing.CompoundLineValues.Single,
                        Alignment = OpenXMLDrawing.PenAlignmentValues.Center
                    },
                    new OpenXMLDrawing.EffectList()
                ),
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
                new OpenXmlChart.ChartShapeProperties(
                    new OpenXMLDrawing.NoFill(),
                    new OpenXMLDrawing.Outline(
                        new OpenXMLDrawing.NoFill()
                    ),
                    new OpenXMLDrawing.EffectList()
                ),
                new OpenXmlChart.CrossingAxis() { Val = Convert.ToUInt32(axisID[0]) },
                new OpenXmlChart.Crosses() { Val = OpenXmlChart.CrossesValues.AutoZero },
                new OpenXmlChart.CrossBetween() { Val = OpenXmlChart.CrossBetweenValues.Between }
            );
            plotAreaElement.AppendChild(valAxis);

            this.Chart.ChartXml.AppendChild(plotAreaElement);
        }
    }
}
