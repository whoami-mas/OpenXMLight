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
using System.Linq;

namespace OpenXMLight
{
    public class WordDocument : IDisposable
    {
        private table.Table tables;

        private WordprocessingDocument? WordProc { get; set; }
        private Document? Doc { get; set; }
        public List<table.Table> Tables => Doc.Body.Elements<Table>().Select(tbl => new table.Table(tbl)).ToList();


        #region Subject facade
        public SettingsPageWord SettingsDocument { get; protected set; }
        private StylesWord StylesDocument { get; set; }
        #endregion

        #region Dispose
        public void Dispose()
        {
            WordProc?.Dispose();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion

        public void Save()
        {
            WordProc?.MainDocumentPart?.Document.Save();
            WordProc?.Dispose();
        }
        public WordDocument(string path, bool overwrite = false)
        {
            if (overwrite)
                File.Delete(path);

            WordProc = File.Exists(path) ? WordprocessingDocument.Open(path, true)
                                                 : WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);

            if (WordProc.MainDocumentPart == null)
                WordProc.AddMainDocumentPart().Document = new Document(new Body());

            Doc = WordProc.MainDocumentPart?.Document;

            SettingsDocument = new();
            StylesDocument = new();
            
            SettingsDocument.GenerateDocumentSettings(Doc);
            StylesDocument.GenerateStyles(WordProc.MainDocumentPart);
        }

        public void AddText(elements.Text text)
        {
            ValidationDocument.ValidationWord(WordProc);

            Doc.Body.AppendChild(text.Properties.Paragraph);
        }
    
        public void AddTable(table.Table table)
        {
            const int TWIPSINPIXELS = 15;

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
                        row.Cells[i].Width = table.Grid.ColumnWidth[i] * TWIPSINPIXELS;


            Doc.Body.AppendChild(table.TableXml);
        }

        public void BuildChart(charts.ChartBuilder chartBuilder)
        {
            chartBuilder.GeneratedTitle();
            chartBuilder.GeneratedAutoTitleDeleted();
            chartBuilder.GeneratedPlotArea();
            chartBuilder.GeneratedLegend();
            chartBuilder.GeneratedPlotVisibleOnly();

            ChartPart chartPart = Doc.MainDocumentPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = chartBuilder.Chart.ChartSpaceXml;

            Paragraph p = new Paragraph();
            Run r = new Run();

            Drawing drawing = new Drawing();
            Inline inline = new();

            inline.Append(
                new Extent() { Cx = chartBuilder.Chart.WidthLong, Cy = chartBuilder.Chart.HeightLong },
                new EffectExtent() { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new DocProperties() { Name = "Диаграмма 1", Id = 1 },
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

        public elements.Endnote AddEndnote(string content, TextProperties? textProp = default)
        {
            textProp ??= new TextProperties();
            elements.Endnote endnote = new elements.Endnote(content, StylesDocument.CreateGetEndnoteStyle(), textProp);

            EndnotesPart endnotesPart = Doc.MainDocumentPart.EndnotesPart ?? Doc.MainDocumentPart.AddNewPart<EndnotesPart>();
            endnotesPart.Endnotes ??= new Endnotes();

            int idEndnote = endnotesPart.Endnotes.Count() == 0 ? 0 : +1;
            endnote.SetID(idEndnote);
            Endnote endnoteXml = new Endnote() { Id = idEndnote };

            endnoteXml.AppendChild(endnote.Properties.Paragraph);
            endnotesPart.Endnotes.AppendChild(endnoteXml);
            endnotesPart.Endnotes.Save();

            return endnote;
        }
    }
}
