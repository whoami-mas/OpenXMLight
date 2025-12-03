using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXMLDrawing = DocumentFormat.OpenXml.Drawing;
using OpenXmlChart = DocumentFormat.OpenXml.Drawing.Charts;
using System.Globalization;

namespace OpenXMLight.Configurations.Elements.Charts
{
    public class PieChart : ChartBuilder
    {
        private ChartData chartData => this.Chart.Data[0];

        public override ChartBuilder SetData(List<ChartData> data)
        {
            if (data.Where(w => w.Labels.Length > 1 || w.Data.Length > 1).ToList().Count > 0)
                throw new Exception("Круговая диаграмма не допускает несколько значений для 1 серии");
            //if (data.Count > 1)
            //    throw new Exception("");
            
            return base.SetData(data);
        }

        internal override void GeneratedTitle()
        {
            OpenXmlChart.Title titleElement = new OpenXmlChart.Title(
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
                                    FontSize = 1400,
                                    Bold = false,
                                    Italic = false,
                                    Underline = OpenXMLDrawing.TextUnderlineValues.None,
                                    Strike = OpenXMLDrawing.TextStrikeValues.NoStrike,
                                    Kerning = 1200,
                                    Spacing = 0,
                                    Baseline = 0
                                }
                        )
                    )
                    )
                );

            this.Chart.ChartXml.Append(titleElement);
        }
        internal override void GeneratedPlotArea()
        {
            OpenXmlChart.PlotArea plotAreaElement = new OpenXmlChart.PlotArea(
                new OpenXmlChart.Layout()
            );

            OpenXmlChart.PieChart pieChart = new OpenXmlChart.PieChart(
                new OpenXmlChart.VaryColors() { Val = true}
            );

            //Title chart
            OpenXmlChart.PieChartSeries pieChartSeries = new OpenXmlChart.PieChartSeries(
                new OpenXmlChart.Index() { Val = 0U },
                new OpenXmlChart.Order() { Val = 0U },
                new OpenXmlChart.SeriesText(
                    new OpenXmlChart.StringReference(
                        new OpenXmlChart.Formula() { Text = "Лист1!$B$1" },
                        new OpenXmlChart.StringCache(
                            new OpenXmlChart.PointCount() { Val = 1U},
                            new OpenXmlChart.StringPoint(
                                new OpenXmlChart.NumericValue() { Text = base.Chart.Title }
                            ) { Index = 0U}
                        )
                    )
                )
            );

            #region Inizialization CategoryAxisData
            OpenXmlChart.CategoryAxisData catAxisData = new OpenXmlChart.CategoryAxisData();
            OpenXmlChart.StringReference strReference = new OpenXmlChart.StringReference(
                new OpenXmlChart.Formula()
                {
                    Text = base.Chart.Data.Count > 1 ? $"Лист1!$A$2:$A${base.Chart.Data.Count + 1}"
                                                                                 : "Лист1!$A$2"
                }
            );
            OpenXmlChart.StringCache strCache = new OpenXmlChart.StringCache(
                new OpenXmlChart.PointCount() { Val = Convert.ToUInt32(base.Chart.Data.Count) }
                );
            #endregion

            #region Inizialization Values
            OpenXmlChart.Values values = new OpenXmlChart.Values();
            OpenXmlChart.NumberReference numReference = new OpenXmlChart.NumberReference(
                new OpenXmlChart.Formula()
                {
                    Text = base.Chart.Data.Count > 1 ? $"Лист1!$A$2:$A${base.Chart.Data.Count + 1}"
                                                                                 : "Лист1!$A$2"
                }
            );
            OpenXmlChart.NumberingCache numCache = new OpenXmlChart.NumberingCache(
                new OpenXmlChart.FormatCode() { Text = "Genereal" },
                new OpenXmlChart.PointCount() { Val = Convert.ToUInt32(base.Chart.Data.Count) }
            );
            #endregion

            for (int i = 0; i < base.Chart.Data.Count; i++)
            {
                OpenXmlChart.DataPoint dataPoint = new OpenXmlChart.DataPoint(
                    new OpenXmlChart.Index() { Val = Convert.ToUInt32(i) },
                    new OpenXmlChart.Bubble3D() { Val = false },
                    new OpenXmlChart.ChartShapeProperties(
                        new OpenXMLDrawing.SolidFill(
                            new OpenXMLDrawing.SchemeColor() { Val = StyleLine[i] }
                        ),
                        new OpenXMLDrawing.Outline(
                            new OpenXMLDrawing.SolidFill(
                                new OpenXMLDrawing.SchemeColor() { Val = OpenXMLDrawing.SchemeColorValues.Light1 }
                            )
                        )
                        { Width = 19050 },
                        new OpenXMLDrawing.EffectList()
                    )
                );

                pieChartSeries.AppendChild(dataPoint);

                #region CategoryAxisData
                for (int j = 0; j < base.Chart.Data[i].Labels.Length; j++)
                {
                    OpenXmlChart.StringPoint strPoint = new OpenXmlChart.StringPoint(
                        new OpenXmlChart.NumericValue() { Text = base.Chart.Data[i].Title }
                    )
                    { Index = Convert.ToUInt32(i) };

                    strCache.AppendChild(strPoint);
                }
                #endregion

                #region Values
                for (int e = 0; e < base.Chart.Data[i].Data.Length; e++)
                {
                    OpenXmlChart.NumericPoint numPoint = new OpenXmlChart.NumericPoint(
                        new OpenXmlChart.NumericValue() { Text = base.Chart.Data[i].Data[e].ToString(CultureInfo.InvariantCulture) }
                        )
                    { Index = Convert.ToUInt32(i) };

                    numCache.AppendChild(numPoint);
                }
                #endregion
            }
            //Append CategoryAxisData
            strReference.AppendChild(strCache);
            catAxisData.AppendChild(strReference);
            pieChartSeries.AppendChild(catAxisData);

            //Append values
            numReference.AppendChild(numCache);
            values.AppendChild(numReference);
            pieChartSeries.AppendChild(values);
            
            pieChart.AppendChild(pieChartSeries);

            plotAreaElement.AppendChild(pieChart);

            this.Chart.ChartXml.AppendChild(plotAreaElement);
        }
    }
}
