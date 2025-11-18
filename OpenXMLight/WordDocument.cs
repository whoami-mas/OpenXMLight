using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLight.Configurations;
using elements = OpenXMLight.Configurations.Elements;
using table = OpenXMLight.Configurations.Elements.Table;
using OpenXMLight.Configurations.Formatting;
using OpenXMLight.SettingsW;
using OpenXMLight.validations;
using charts = OpenXMLight.Configurations.Elements.Charts;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using OpenXmlDrawing = DocumentFormat.OpenXml.Drawing;
using OpenXmlChart = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLight
{
    public class WordDocument : IDisposable
    {
        private WordprocessingDocument? WordProc { get; set; } = null;
        private Document? Doc { get; set; } = null;
        private SettingsPageWord? SettingsDocument { get; set; }

        #region Dispose
        public void Dispose()
        {
            WordProc?.Dispose();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion

        public WordDocument(string pathDocument, bool overwrite = false)
        {
            WordProc = File.Exists(pathDocument) ? WordprocessingDocument.Open(pathDocument, true)
                                                     : WordprocessingDocument.Create(pathDocument, WordprocessingDocumentType.Document);

            MainDocumentPart mainPart;

            if (WordProc.MainDocumentPart != null)
            {
                if (overwrite)
                {
                    WordProc.DeletePart(WordProc.MainDocumentPart);
                    WordProc.AddMainDocumentPart().Document = new Document(new Body());
                }

                mainPart = WordProc.MainDocumentPart;
            }
            else
                mainPart = WordProc.AddMainDocumentPart();

            Doc = mainPart.Document;

            SettingsDocument = new();
            SettingsDocument.GenerateDocumentSettings(Doc);
        }

        public void AddText(elements.Text text)
        {
            ValidationDocument.ValidationWord(WordProc);

            Doc.Body.AppendChild(text.Properties.Paragraph);
        }
    
        public void AddTable(table.Table table)
        {
            if (table.Grid?.ColumnWidth == null)
            {
                int? maxCountCell = table.Rows?.Select(s => s?.Cells?.Count).DefaultIfEmpty(null).Max();
                string widthColumn = ((SettingsDocument?.WidthPage - SettingsDocument?.MarginLeft - SettingsDocument?.MarginRight) / maxCountCell).ToString();

                for (int i = 0; i < maxCountCell; i++)
                    table.TableXml.Elements<TableGrid>().First().AppendChild(
                        new GridColumn() { Width = widthColumn }
                    );

                foreach (var row in table.Rows)
                    foreach (var cell in row.Cells)
                        cell.Width = int.Parse(widthColumn);
            }
            else
                foreach (var row in table.Rows)
                    for (int i = 0; i < table.Grid.ColumnWidth.Length; i++)
                        row.Cells[i].Width = table.Grid.ColumnWidth[i];


            Doc.Body.AppendChild(table.TableXml);
        }

        public void AddChart(charts.ColumnChart columnChart)
        {
            ChartPart chartPart = Doc.MainDocumentPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = columnChart.chartSpace;

            Paragraph p = new Paragraph();
            Run r = new Run();

            Drawing drawing = new Drawing();
            Inline inline = new();

            inline.Append(
                new Extent() { Cx = 5486400L, Cy = 3200400L },
                new EffectExtent() { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new DocProperties() { Name = "Диаграмма 1", Id = 1},
                new NonVisualGraphicFrameDrawingProperties()
                );
            inline.AppendChild(
                new OpenXmlDrawing.Graphic(
                    new OpenXmlDrawing.GraphicData(
                        new OpenXmlChart.ChartReference()
                        {
                            Id = Doc.MainDocumentPart.GetIdOfPart(chartPart)
                        }
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                )
            );
            drawing.AppendChild(inline);

            r.AppendChild(drawing);
            p.AppendChild(r);

            Doc.Body.AppendChild(p);
        }
    }
}
