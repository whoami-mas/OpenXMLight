using ConsoleApp1;
using OpenXMLight;
using OpenXMLight.Configurations;
using OpenXMLight.Configurations.Elements;
using OpenXMLight.Configurations.Elements.TableElements;
using OpenXMLight.Configurations.Elements.TableElements.Formattings.MarginComponents;
using OpenXMLight.Configurations.Elements.TableElements.Models;
using OpenXMLight.Configurations.Formatting;

try
{
    //string path = @"C:\Users\bushk\Desktop\Reportings risks\act_rep_4_2025.xlsx";
    //string path = @"testingTable.docx";
    string path = @"F:\тестовые проекты\ConsoleApp1\тест\testing.docx";

    using (WordDocument document = new WordDocument(path, true))
    {
        Table tableSign = document.AddTable()
                        .SetWidth(
                            w =>
                            {
                                w.Width = "11,5";
                                w.Type = TypeWidthTable.Cm;
                            }
                        )
                        .SetBorders(
                        b=>
                        {
                            b.LineWidth = 1;
                            b.LineType = BordersType.Single;
                        }
                        )
                        .IsFixed(true)
                        .AddRows(
                            r => r
                            .AddCell(
                                    c => c
                                    .AddParagraph(
                                        p => p.SetRun(
                                            new RunBuilder()
                                            .SetText("За отчетный период")
                                            .SetFontSize(9)
                                            .SetFontFamily(FontsFamily.TimesNewRoman)
                                            .SetBold(true)
                                            )
                                        .SetSpacingBetweenLines(new SpacingBetweenLines()
                                        {
                                            After = 0,
                                            Before = 0,
                                            Line = 250
                                        })
                                        )
                                    .SetWidth(
                                        w =>
                                        {
                                            w.Width = "3,25";
                                            w.Type = TypeWidthTable.Cm;
                                        }
                                        )
                                )
                            .AddCell(
                                    c => c
                                    .AddParagraph(
                                        p => p.SetRun(
                                            new RunBuilder()
                                            .SetText("Fixed 10 event risks")
                                            .SetFontSize(9)
                                            .SetFontFamily(FontsFamily.TimesNewRoman)
                                            )
                                        .SetSpacingBetweenLines(new SpacingBetweenLines()
                                        {
                                            After = 0,
                                            Before = 0,
                                            Line = 250
                                        })
                                        )
                                    .SetWidth(
                                        w =>
                                        {
                                            w.Width = "2,25";
                                            w.Type = TypeWidthTable.Cm;
                                        }
                                    )
                                )
                            .AddCell(
                                    c => c
                                    .AddParagraph(
                                        p => p.SetRun(
                                            new RunBuilder()
                                            .SetText("")
                                            .SetFontSize(9)
                                            .SetFontFamily(FontsFamily.TimesNewRoman)
                                            )
                                        .SetSpacingBetweenLines(new SpacingBetweenLines()
                                        {
                                            After = 0,
                                            Before = 0,
                                            Line = 250
                                        })
                                        )
                                    .SetWidth(
                                        w =>
                                        {
                                            w.Width = "2,25";
                                            w.Type = TypeWidthTable.Cm;
                                        }
                                        )
                                )
                            .AddCell(
                                    c => c
                                    .AddParagraph(
                                        p => p.SetRun(
                                            new RunBuilder()
                                            .SetText("")
                                            .SetFontSize(9)
                                            .SetFontFamily(FontsFamily.TimesNewRoman)
                                            )
                                        .SetSpacingBetweenLines(new SpacingBetweenLines()
                                        {
                                            After = 0,
                                            Before = 0,
                                            Line = 250
                                        })
                                        )
                                    .SetWidth(
                                        w =>
                                        {
                                            w.Width = "3,75";
                                            w.Type = TypeWidthTable.Cm;
                                        }
                                        )
                                )
                        )
                        .AddRows(
                            r => r
                            .AddCell(
                                    c => c
                                    .AddParagraph()
                                    .SetWidth(
                                        w =>
                                        {
                                            w.Width = "3,25";
                                            w.Type = TypeWidthTable.Cm;
                                        }
                                        )
                                )
                            .AddCell(
                                    c => c
                                    .AddParagraph(
                                        )
                                )
                            .AddCell(
                                    c => c
                                    .AddParagraph(
                                        p => p.SetRun(
                                            new RunBuilder()
                                            .SetText("123")
                                            .SetFontSize(9)
                                            .SetFontFamily(FontsFamily.TimesNewRoman)
                                            .SetBold(true)
                                            ))
                                )
                            .AddCell(
                                    c => c
                                    .AddParagraph()
                                )
                        )
                        .AddRows(
                            r => r
                            .AddCell(
                                    c => c
                                    .AddParagraph(
                                        p => p.SetRun(
                                            new RunBuilder()
                                            .SetText("546")
                                            .SetFontSize(9)
                                            .SetFontFamily(FontsFamily.TimesNewRoman)
                                            )
                                        .SetSpacingBetweenLines(new SpacingBetweenLines()
                                        {
                                            After = 0,
                                            Before = 0,
                                            Line = 250
                                        })
                                        )
                                    .SetColor(Color.FromHex("#00ef80"))
                                    .SetWidth(
                                        w =>
                                        {
                                            w.Width = "3,25";
                                            w.Type = TypeWidthTable.Cm;
                                        }
                                        )
                                )
                            .AddCell(
                                    c => c
                                    .AddParagraph(
                                        p => p.SetRun(
                                            new RunBuilder()
                                            .SetText("(должность)")
                                            .SetFontSize(9)
                                            .SetFontFamily(FontsFamily.TimesNewRoman)
                                            )
                                        .SetSpacingBetweenLines(new SpacingBetweenLines()
                                        {
                                            After = 0,
                                            Before = 0,
                                            Line = 250
                                        })
                                        .SetAlignment(HorizontalAlignments.Center)
                                        )
                                    .SetWidth(
                                        w =>
                                        {
                                            w.Width = "2,25";
                                            w.Type = TypeWidthTable.Cm;
                                        }
                                        )
                                )
                            .AddCell(
                                    c => c
                                    .AddParagraph(
                                        p => p.SetRun(
                                            new RunBuilder()
                                            .SetText("(подпись)")
                                            .SetFontSize(9)
                                            .SetFontFamily(FontsFamily.TimesNewRoman)
                                            )
                                        .SetSpacingBetweenLines(new SpacingBetweenLines()
                                        {
                                            After = 0,
                                            Before = 0,
                                            Line = 250
                                        })
                                        .SetAlignment(HorizontalAlignments.Center)
                                        )
                                    .SetWidth(
                                        w =>
                                        {
                                            w.Width = "2,25";
                                            w.Type = TypeWidthTable.Cm;
                                        }
                                        )
                                )
                            .AddCell(
                                    c => c
                                    .AddParagraph(
                                        p => p.SetRun(
                                            new RunBuilder()
                                            .SetText("(расшифровка подписи)")
                                            .SetFontSize(9)
                                            .SetFontFamily(FontsFamily.TimesNewRoman)
                                            )
                                        .SetSpacingBetweenLines(new SpacingBetweenLines()
                                        {
                                            After = 0,
                                            Before = 0,
                                            Line = 250
                                        })
                                        .SetAlignment(HorizontalAlignments.Center)
                                        )
                                    .SetWidth(
                                        w =>
                                        {
                                            w.Width = "3,75";
                                            w.Type = TypeWidthTable.Cm;
                                        }
                                        )
                                )
                        )
                        .SetMargin("0,15", "0", "0,15", "0");

    }
}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}
