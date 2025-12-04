using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXMLDrawing = DocumentFormat.OpenXml.Drawing;
using OpenXmlChart = DocumentFormat.OpenXml.Drawing.Charts;
using System.Globalization;
using OpenXMLight.Spreadsheet.Formatting;

namespace OpenXMLight.Configurations.Elements.Charts
{
    public class LineChart : ChartBuilder
    {
        internal bool isAxisRight = false;

        public LineChart()
        {
            AxesID = new()
            {
                {Axes.Left, Random.Shared.Next(100000000, 999999999) },
                {Axes.Bottom, Random.Shared.Next(100000000, 999999999) }
            };
        }

        internal override void GeneratedPlotArea()
        {
            if (Chart.Data.Where(w => w.orientationY == Orientation.Right).Count() > 0)
            {
                AxesID.Add(Axes.Right, Random.Shared.Next(100000000, 999999999));
                isAxisRight = true;
            }

            OpenXmlChart.PlotArea plotAreaElement = new OpenXmlChart.PlotArea(
                new OpenXmlChart.Layout()
                );

            int indexSer = 0;

            #region Append left
            //LineChart
            OpenXmlChart.LineChart lineChartLeft = new OpenXmlChart.LineChart(
                new OpenXmlChart.Grouping() { Val = OpenXmlChart.GroupingValues.Standard },
                new OpenXmlChart.VaryColors() { Val = false }
                );

            List<ChartData> dataChartLeft = this.Chart.Data.Where(w => w.orientationY == Orientation.Left).ToList();
            foreach (ChartData ser in dataChartLeft)
            {
                //Index
                OpenXmlChart.LineChartSeries lineSeries = new OpenXmlChart.LineChartSeries(
                    new OpenXmlChart.Index() { Val = Convert.ToUInt32(indexSer) },
                    new OpenXmlChart.Order() { Val = Convert.ToUInt32(indexSer) }
                );

                //SeriesText
                lineSeries.AppendChild(
                    new OpenXmlChart.SeriesText(
                        new OpenXmlChart.StringReference(
                            new OpenXmlChart.Formula() { Text = $"Лист1!${(char)('B' + indexSer)}$1" },
                            new OpenXmlChart.StringCache(
                                new OpenXmlChart.PointCount() { Val = 1U },
                                new OpenXmlChart.StringPoint(
                                    new OpenXmlChart.NumericValue() { Text = ser.Title }
                                    )
                                { Index = 0U }
                                )
                            )
                        )
                );

                //ShapeProperties
                lineSeries.AppendChild(
                    new OpenXmlChart.ChartShapeProperties(
                            new OpenXMLDrawing.Outline(
                                new OpenXMLDrawing.SolidFill(
                                    new OpenXMLDrawing.SchemeColor() { Val = this.StyleLine[indexSer] }
                                    )
                                )
                            { Width = 35000, CapType = OpenXMLDrawing.LineCapValues.Round }, //28575
                            new OpenXMLDrawing.EffectList()
                    )
                );

                //Marker
                lineSeries.AppendChild(
                    new OpenXmlChart.Marker(
                        new OpenXmlChart.Symbol() { Val = OpenXmlChart.MarkerStyleValues.None }
                    )
                );

                #region CategoryAxisDate
                OpenXmlChart.CategoryAxisData categoryAxisDate = new OpenXmlChart.CategoryAxisData();
                OpenXmlChart.StringReference strReference = new OpenXmlChart.StringReference(
                    new OpenXmlChart.Formula()
                    {
                        Text = ser.Labels.Length > 1 ? $"Лист1!$A$2:$A${ser.Labels.Length}"
                                                                     : "Лист1!$A$2"
                    }
                );
                OpenXmlChart.StringCache strCache = new OpenXmlChart.StringCache(
                    new OpenXmlChart.PointCount() { Val = Convert.ToUInt32(ser.Labels.Length) }
                );
                for (int j = 0; j < ser.Data.Length; j++)
                    strCache.AppendChild(
                        new OpenXmlChart.StringPoint(
                            new OpenXmlChart.NumericValue() { Text = ser.Labels[j] }
                            )
                        { Index = Convert.ToUInt32(j) }
                    );
                strReference.AppendChild(strCache);
                categoryAxisDate.AppendChild(strReference);
                lineSeries.AppendChild(categoryAxisDate);
                #endregion

                #region Values
                OpenXmlChart.Values values = new OpenXmlChart.Values();
                OpenXmlChart.NumberReference numberReference = new OpenXmlChart.NumberReference(
                    new OpenXmlChart.Formula()
                    {
                        Text = ser.Data.Length > 1 ? $"Лист1!${(char)('B' + indexSer)}$2:${(char)('B' + indexSer)}${ser.Data.Length}"
                                                                   : $"Лист1!${(char)('B' + indexSer)}$"
                    }
                    );
                OpenXmlChart.NumberingCache numCache = new OpenXmlChart.NumberingCache(
                    new OpenXmlChart.FormatCode() { Text = "General" },
                    new OpenXmlChart.PointCount() { Val = Convert.ToUInt32(ser.Data.Length) }
                    );
                for (int j = 0; j < ser.Data.Length; j++)
                    numCache.AppendChild(
                        new OpenXmlChart.NumericPoint(
                            new OpenXmlChart.NumericValue() { Text = ser.Data[j].ToString(CultureInfo.InvariantCulture) }
                            )
                        { Index = Convert.ToUInt32(j) }
                        );
                numberReference.AppendChild(numCache);
                values.AppendChild(numberReference);

                lineSeries.AppendChild(values);
                #endregion

                lineChartLeft.AppendChild(lineSeries);

                indexSer++;
            }
            //for (int i = indexSer; i < dataChartLeft.Count; i++)
            //{
            //    //Index
            //    OpenXmlChart.LineChartSeries lineSeries = new OpenXmlChart.LineChartSeries(
            //        new OpenXmlChart.Index() { Val = Convert.ToUInt32(i) },
            //        new OpenXmlChart.Order() { Val = Convert.ToUInt32(i) }
            //        );

            //    //SeriesText
            //    lineSeries.AppendChild(
            //        new OpenXmlChart.SeriesText(
            //            new OpenXmlChart.StringReference(
            //                new OpenXmlChart.Formula() { Text = $"Лист1!${(char)('B' + i)}$1" },
            //                new OpenXmlChart.StringCache(
            //                    new OpenXmlChart.PointCount() { Val = 1U },
            //                    new OpenXmlChart.StringPoint(
            //                        new OpenXmlChart.NumericValue() { Text = this.Chart.Data[i].Title }
            //                        )
            //                    { Index = 0U }
            //                    )
            //                )
            //            )
            //    );

            //    //ShapeProperties
            //    lineSeries.AppendChild(
            //        new OpenXmlChart.ChartShapeProperties(
            //                new OpenXMLDrawing.Outline(
            //                    new OpenXMLDrawing.SolidFill(
            //                        new OpenXMLDrawing.SchemeColor() { Val = this.StyleLine[i] }
            //                        )
            //                    )
            //                { Width = 35000, CapType = OpenXMLDrawing.LineCapValues.Round}, //28575
            //                new OpenXMLDrawing.EffectList()
            //            )
            //        );
            //    //Marker
            //    lineSeries.AppendChild(
            //        new OpenXmlChart.Marker(
            //            new OpenXmlChart.Symbol() { Val = OpenXmlChart.MarkerStyleValues.None}
            //        )
            //    );

            //    #region CategoryAxisDate
            //    OpenXmlChart.CategoryAxisData categoryAxisDate = new OpenXmlChart.CategoryAxisData();
            //    OpenXmlChart.StringReference strReference = new OpenXmlChart.StringReference(
            //        new OpenXmlChart.Formula()
            //        {
            //            Text = this.Chart.Data[i].Labels.Length > 1 ? $"Лист1!$A$2:$A${this.Chart.Data[i].Labels.Length}"
            //                                                         : "Лист1!$A$2"
            //        }
            //        );
            //    OpenXmlChart.StringCache strCache = new OpenXmlChart.StringCache(
            //        new OpenXmlChart.PointCount() { Val = Convert.ToUInt32(this.Chart.Data[i].Labels.Length) }
            //        );
            //    for (int j = 0; j < this.Chart.Data[i].Data.Length; j++)
            //        strCache.AppendChild(
            //            new OpenXmlChart.StringPoint(
            //                new OpenXmlChart.NumericValue() { Text = this.Chart.Data[i].Labels[j] }
            //                )
            //            { Index = Convert.ToUInt32(j) }
            //        );
            //    strReference.AppendChild(strCache);
            //    categoryAxisDate.AppendChild(strReference);
            //    lineSeries.AppendChild(categoryAxisDate);
            //    #endregion

            //    #region Values
            //    OpenXmlChart.Values values = new OpenXmlChart.Values();
            //    OpenXmlChart.NumberReference numberReference = new OpenXmlChart.NumberReference(
            //        new OpenXmlChart.Formula()
            //        {
            //            Text = this.Chart.Data[i].Data.Length > 1 ? $"Лист1!${(char)('B' + i)}$2:${(char)('B' + i)}${this.Chart.Data[i].Data.Length}"
            //                                                       : $"Лист1!${(char)('B' + i)}$"
            //        }
            //        );
            //    OpenXmlChart.NumberingCache numCache = new OpenXmlChart.NumberingCache(
            //        new OpenXmlChart.FormatCode() { Text = "General" },
            //        new OpenXmlChart.PointCount() { Val = Convert.ToUInt32(this.Chart.Data[i].Data.Length) }
            //        );
            //    for (int j = 0; j < this.Chart.Data[i].Data.Length; j++)
            //        numCache.AppendChild(
            //            new OpenXmlChart.NumericPoint(
            //                new OpenXmlChart.NumericValue() { Text = this.Chart.Data[i].Data[j].ToString(CultureInfo.InvariantCulture) }
            //                )
            //            { Index = Convert.ToUInt32(j) }
            //            );
            //    numberReference.AppendChild(numCache);
            //    values.AppendChild(numberReference);

            //    lineSeries.AppendChild(values);
            //    #endregion

            //    lineChart.AppendChild(lineSeries);

            //    indexSer = i;
            //}

            lineChartLeft.Append(
               new OpenXmlChart.AxisId() { Val = Convert.ToUInt32(AxesID[Axes.Bottom]) },
               new OpenXmlChart.AxisId() { Val = Convert.ToUInt32(AxesID[Axes.Left]) }
            );

            plotAreaElement.AppendChild(lineChartLeft);
            #endregion

            #region Append right
            if (base.IsRightAxis)
            {
                //LineChart
                OpenXmlChart.LineChart lineChartRight = new OpenXmlChart.LineChart(
                    new OpenXmlChart.Grouping() { Val = OpenXmlChart.GroupingValues.Standard },
                    new OpenXmlChart.VaryColors() { Val = false }
                    );

                List<ChartData> dataChartRight = this.Chart.Data.Where(w => w.orientationY == Orientation.Right).ToList();
                foreach (ChartData ser in dataChartRight)
                {
                    //Index
                    OpenXmlChart.LineChartSeries lineSeries = new OpenXmlChart.LineChartSeries(
                        new OpenXmlChart.Index() { Val = Convert.ToUInt32(indexSer) },
                        new OpenXmlChart.Order() { Val = Convert.ToUInt32(indexSer) }
                    );

                    //SeriesText
                    lineSeries.AppendChild(
                        new OpenXmlChart.SeriesText(
                            new OpenXmlChart.StringReference(
                                new OpenXmlChart.Formula() { Text = $"Лист1!${(char)('B' + indexSer)}$1" },
                                new OpenXmlChart.StringCache(
                                    new OpenXmlChart.PointCount() { Val = 1U },
                                    new OpenXmlChart.StringPoint(
                                        new OpenXmlChart.NumericValue() { Text = ser.Title }
                                        )
                                    { Index = 0U }
                                    )
                                )
                            )
                    );

                    //ShapeProperties
                    lineSeries.AppendChild(
                        new OpenXmlChart.ChartShapeProperties(
                                new OpenXMLDrawing.Outline(
                                    new OpenXMLDrawing.SolidFill(
                                        new OpenXMLDrawing.SchemeColor() { Val = this.StyleLine[indexSer] }
                                        )
                                    )
                                { Width = 35000, CapType = OpenXMLDrawing.LineCapValues.Round }, //28575
                                new OpenXMLDrawing.EffectList()
                        )
                    );

                    //Marker
                    lineSeries.AppendChild(
                        new OpenXmlChart.Marker(
                            new OpenXmlChart.Symbol() { Val = OpenXmlChart.MarkerStyleValues.None }
                        )
                    );

                    #region CategoryAxisDate
                    OpenXmlChart.CategoryAxisData categoryAxisDate = new OpenXmlChart.CategoryAxisData();
                    OpenXmlChart.StringReference strReference = new OpenXmlChart.StringReference(
                        new OpenXmlChart.Formula()
                        {
                            Text = ser.Labels.Length > 1 ? $"Лист1!$A$2:$A${ser.Labels.Length}"
                                                                         : "Лист1!$A$2"
                        }
                    );
                    OpenXmlChart.StringCache strCache = new OpenXmlChart.StringCache(
                        new OpenXmlChart.PointCount() { Val = Convert.ToUInt32(ser.Labels.Length) }
                    );
                    for (int j = 0; j < ser.Data.Length; j++)
                        strCache.AppendChild(
                            new OpenXmlChart.StringPoint(
                                new OpenXmlChart.NumericValue() { Text = ser.Labels[j] }
                                )
                            { Index = Convert.ToUInt32(j) }
                        );
                    strReference.AppendChild(strCache);
                    categoryAxisDate.AppendChild(strReference);
                    lineSeries.AppendChild(categoryAxisDate);
                    #endregion

                    #region Values
                    OpenXmlChart.Values values = new OpenXmlChart.Values();
                    OpenXmlChart.NumberReference numberReference = new OpenXmlChart.NumberReference(
                        new OpenXmlChart.Formula()
                        {
                            Text = ser.Data.Length > 1 ? $"Лист1!${(char)('B' + indexSer)}$2:${(char)('B' + indexSer)}${ser.Data.Length}"
                                                                       : $"Лист1!${(char)('B' + indexSer)}$"
                        }
                    );
                    OpenXmlChart.NumberingCache numCache = new OpenXmlChart.NumberingCache(
                        new OpenXmlChart.FormatCode() { Text = "General" },
                        new OpenXmlChart.PointCount() { Val = Convert.ToUInt32(ser.Data.Length) }
                    );
                    for (int j = 0; j < ser.Data.Length; j++)
                        numCache.AppendChild(
                            new OpenXmlChart.NumericPoint(
                                new OpenXmlChart.NumericValue()
                                {
                                    Text = ser.TypeValueSeries switch
                                    {
                                        TypeSeries.Percent => (ser.Data[j] * 0.01).ToString(CultureInfo.InvariantCulture),
                                        TypeSeries.General => ser.Data[j].ToString(CultureInfo.InvariantCulture),
                                        _ => throw new ArgumentException("Неизвестный тип серии")
                                    }
                                }
                            )
                            { Index = Convert.ToUInt32(j) }
                        );
                    numberReference.AppendChild(numCache);
                    values.AppendChild(numberReference);

                    lineSeries.AppendChild(values);
                    #endregion

                    lineChartRight.AppendChild(lineSeries);

                    indexSer++;
                }

                lineChartRight.Append(
                   new OpenXmlChart.AxisId() { Val = Convert.ToUInt32(AxesID[Axes.Bottom]) },
                   new OpenXmlChart.AxisId() { Val = Convert.ToUInt32(AxesID[Axes.Right]) }
                );

                plotAreaElement.AppendChild(lineChartRight);
            }
            #endregion

            //CategoryAxis
            OpenXmlChart.CategoryAxis catAxis = new OpenXmlChart.CategoryAxis(
                new OpenXmlChart.AxisId() { Val = Convert.ToUInt32(AxesID[Axes.Bottom]) },
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
                                new OpenXMLDrawing.LuminanceModulation() { Val = 15000},
                                new OpenXMLDrawing.LuminanceOffset() { Val = 85000 }
                            ) { Val = OpenXMLDrawing.SchemeColorValues.Text1}
                        )
                    ) { Width = 9525,
                            CapType = OpenXMLDrawing.LineCapValues.Flat,
                            CompoundLineType = OpenXMLDrawing.CompoundLineValues.Single,
                            Alignment = OpenXMLDrawing.PenAlignmentValues.Center},
                    new OpenXMLDrawing.EffectList()
                ),
                new OpenXmlChart.CrossingAxis() { Val = Convert.ToUInt32(AxesID[Axes.Left]) },
                new OpenXmlChart.Crosses() { Val = OpenXmlChart.CrossesValues.AutoZero },
                new OpenXmlChart.AutoLabeled() { Val = true },
                new OpenXmlChart.LabelAlignment() { Val = OpenXmlChart.LabelAlignmentValues.Center },
                new OpenXmlChart.LabelOffset() { Val = Convert.ToUInt16(100) },
                new OpenXmlChart.NoMultiLevelLabels() { Val = false }
            );
            plotAreaElement.AppendChild(catAxis);

            //ValueAxis
            OpenXmlChart.ValueAxis valAxis = new OpenXmlChart.ValueAxis(
                new OpenXmlChart.AxisId() { Val = Convert.ToUInt32(AxesID[Axes.Left]) },
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
                new OpenXmlChart.CrossingAxis() { Val = Convert.ToUInt32(AxesID[Axes.Bottom]) },
                new OpenXmlChart.Crosses() { Val = OpenXmlChart.CrossesValues.AutoZero },
                new OpenXmlChart.CrossBetween() { Val = OpenXmlChart.CrossBetweenValues.Between }
            );
            plotAreaElement.AppendChild(valAxis);

            if(base.IsRightAxis)
            {
                //ValueAxis
                OpenXmlChart.ValueAxis valAxisRight = new OpenXmlChart.ValueAxis(
                    new OpenXmlChart.AxisId() { Val = Convert.ToUInt32(AxesID[Axes.Right]) },
                    new OpenXmlChart.Scaling(
                        new OpenXmlChart.Orientation() { Val = OpenXmlChart.OrientationValues.MinMax }
                    ),
                    new OpenXmlChart.Delete() { Val = false },
                    new OpenXmlChart.AxisPosition() { Val = OpenXmlChart.AxisPositionValues.Right },
                    new OpenXmlChart.NumberingFormat() { FormatCode = base.TypeFormatAxis.Value, SourceLinked = false },
                    new OpenXmlChart.MajorTickMark() { Val = OpenXmlChart.TickMarkValues.Outside },
                    new OpenXmlChart.MinorTickMark() { Val = OpenXmlChart.TickMarkValues.None },
                    new OpenXmlChart.TickLabelPosition() { Val = OpenXmlChart.TickLabelPositionValues.NextTo },
                    new OpenXmlChart.ChartShapeProperties(
                        new OpenXMLDrawing.NoFill(),
                        new OpenXMLDrawing.Outline(
                            new OpenXMLDrawing.NoFill()
                        ),
                        new OpenXMLDrawing.EffectList()
                    ),
                    new OpenXmlChart.CrossingAxis() { Val = Convert.ToUInt32(AxesID[Axes.Bottom]) },
                    new OpenXmlChart.Crosses() { Val = OpenXmlChart.CrossesValues.Maximum },
                    new OpenXmlChart.CrossBetween() { Val = OpenXmlChart.CrossBetweenValues.Between }
                );
                plotAreaElement.AppendChild(valAxisRight);
            }

            this.Chart.ChartXml.AppendChild(plotAreaElement);
        }
    }
}
