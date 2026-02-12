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
                                    .AddParagraph(
                                        p => p.SetRun(
                                            new RunBuilder()
                                            .SetText("Ответственное лицо")
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
                                            .SetText("__________")
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
                                            .SetText("__________")
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
                                            .SetText("_______________")
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
                        .AddRows(
                            r => r
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
                        .Merge(0, 1, 0, 3)
                        .SetMargin("0,15", "0", "0,15", "0");

        //for (int i = tableSign.Rows[0].Cells[1].Paragraphs.Count - 1; i != 0; i--)
        //    if (string.IsNullOrWhiteSpace(tableSign.Rows[0].Cells[1].Paragraphs[i].AllText))
        //        tableSign.Rows[0].Cells[1].Paragraphs.Remove(tableSign.Rows[0].Cells[1].Paragraphs[i]);
    }
}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}
