using OpenXMLight.Configurations.Formatting;
using OpenXMLight.validations;
using System.Linq;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlElement = DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlDrawing = DocumentFormat.OpenXml.Drawing;
using OpenXmlDrawingWp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using OpenXmlChart = DocumentFormat.OpenXml.Drawing.Charts;

using OpenXMLight.Configurations.Elements.TableElements;
using OpenXMLight.Configurations.Elements.TableElements.Models;
using OpenXMLight.Configurations.Elements.Charts;
using OpenXMLight.Configurations.Elements;
using OpenXMLight.Configurations.WordContext;
using OpenXMLight.Configurations.Parts;

namespace OpenXMLight
{
    public class WordDocument : IDisposable
    {
        private OpenXmlPackaging.WordprocessingDocument? WordProc { get; set; }
        private OpenXmlElement.Document? Doc { get; set; }



        public ElementCollection<Table> Tables => new(Doc.Body.Elements<OpenXmlElement.Table>().Select(s => new Table(s))) {Parent = Doc.Body };
        public SettingsPageWord SettingsDocument { get; protected set; }

        private Context Context { get; init; }



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

            WordProc = File.Exists(path) ? OpenXmlPackaging.WordprocessingDocument.Open(path, true)
                                                 : OpenXmlPackaging.WordprocessingDocument.Create(path, OpenXml.WordprocessingDocumentType.Document);

            if (WordProc.MainDocumentPart == null)
                WordProc.AddMainDocumentPart().Document = new OpenXmlElement.Document(new OpenXmlElement.Body());

            Doc = WordProc.MainDocumentPart?.Document;

            Context = Context.GetInstance(WordProc.MainDocumentPart); ///TODO Testing

            SettingsDocument = new();
            
            SettingsDocument.GenerateDocumentSettings(Doc);
        }



        public void BreakPage() => Doc.Body.AppendChild(new OpenXmlElement.Paragraph(new OpenXmlElement.Run(new OpenXmlElement.Break() { Type = OpenXmlElement.BreakValues.Page })));
        //Testing new code
        public ParagraphBuilder AddParagraph()
        {
            OpenXmlElement.Paragraph p = new();
            Doc.Body.AppendChild(p);
            return new ParagraphBuilder(p);
        }
        public ParagraphBuilder GetParagraph()
        {
            Doc.Body.Elements<OpenXmlElement.Paragraph>().First();
            return new ParagraphBuilder(Doc.Body.Elements<OpenXmlElement.Paragraph>().First());
        }
            
        public TableBuilder AddTable()
        {
            OpenXmlElement.Table tbl = new();
            Doc.Body.AppendChild(tbl);
            return new TableBuilder(tbl, SettingsDocument);
        }

        public void BuildChart(ChartBuilder chartBuilder)
        {
            chartBuilder.GeneratedTitle();
            chartBuilder.GeneratedAutoTitleDeleted();
            chartBuilder.GeneratedPlotArea();
            chartBuilder.GeneratedLegend();
            chartBuilder.GeneratedPlotVisibleOnly();

            OpenXmlPackaging.ChartPart chartPart = Doc.MainDocumentPart.AddNewPart<OpenXmlPackaging.ChartPart>();
            chartPart.ChartSpace = chartBuilder.Chart.ChartSpaceXml;

            OpenXmlElement.Paragraph p = new OpenXmlElement.Paragraph();
            OpenXmlElement.Run r = new OpenXmlElement.Run();

            OpenXmlElement.Drawing drawing = new OpenXmlElement.Drawing();
            OpenXmlDrawingWp.Inline inline = new();

            inline.Append(
                new OpenXmlDrawingWp.Extent() { Cx = chartBuilder.Chart.WidthLong, Cy = chartBuilder.Chart.HeightLong },
                new OpenXmlDrawingWp.EffectExtent() { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new OpenXmlDrawingWp.DocProperties() { Name = "Диаграмма 1", Id = 1 },
                new OpenXmlDrawingWp.NonVisualGraphicFrameDrawingProperties()
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

        public EndnoteBuilder AddEndnote(string content) => new EndnoteBuilder(Context.Endnotes.AddEndnote(content));
    }
}
