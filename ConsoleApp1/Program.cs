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

    using (WordDocument word = new WordDocument(path, true))
    {
        word.SettingsDocument.MarginTop = 49;
        word.SettingsDocument.MarginLeft = 49;
        word.SettingsDocument.MarginRight = 49;
        word.SettingsDocument.MarginBottom = 49;


        word.AddParagraph().SetRun(
            new RunBuilder()
                .SetText($"Сводный отчет ключевых нефинансовых индикаторов риска за 1 2026 года")
                .SetBold(true)
                .SetFontFamily(FontsFamily.TimesNewRoman)
                .SetFontSize(12)
            )
            .SetAlignment(HorizontalAlignments.Center);

        Table table = word.AddTable()
            .IsFixed(true)
            .SetWidth(w =>
            {
                w.Width = "100";
                w.Type = TypeWidthTable.Pct;
            })
            .SetBorders(
                b => new Borders()
                {
                    LineWidth = 1,
                    LineType = BordersType.Single
                }
            )
            .AddRows(
                r =>
                r.AddCell(
                    c => c.AddParagraph(
                        p => p.SetRun(
                            new RunBuilder()
                                .SetText("Риск")
                                .SetBold(true)
                                .SetFontSize(9)
                                .SetFontFamily(FontsFamily.TimesNewRoman)
                            )
                        .SetAlignment(HorizontalAlignments.Center)
                        )

                    .SetVerticalAlignment(VerticalAlignments.Center)
                    )
                .AddCell(
                    c => c.AddParagraph(
                        p => p.SetRun(
                            new RunBuilder()
                            .SetText("Проверка перевернутого текста")
                            .SetBold(true)
                            .SetFontSize(9)
                            .SetFontFamily(FontsFamily.TimesNewRoman)
                            )
                        .SetAlignment(HorizontalAlignments.Center)
                        )
                    .SetWidth(
                        w =>
                        {
                            w.Width = "12";
                            w.Type = TypeWidthTable.Pct;
                        }
                    )
                    .SetVerticalAlignment(VerticalAlignments.Center)
                    )
                .AddCell(
                    c => c.AddParagraph(
                        p => p.SetRun(
                            new RunBuilder()
                            .SetText("№")
                            .SetBold(true)
                            .SetFontSize(9)
                            .SetFontFamily(FontsFamily.TimesNewRoman)
                            )
                        .SetAlignment(HorizontalAlignments.Center)
                        )

                    .SetVerticalAlignment(VerticalAlignments.Center)
                    )
            );

        for (int i = 1; i <= 3; i++)
        {
            table.Rows.Add(
                new RowBuilder()
                    .AddCell(
                        c => c.AddParagraph(
                            p => p.SetRun(
                                new RunBuilder()
                                .SetText($"Тип_{i}")
                                .SetFontSize(9)
                                .SetFontFamily(FontsFamily.TimesNewRoman)
                                )
                            .SetAlignment(HorizontalAlignments.Center)
                            )

                    )
                    .AddCell(
                        c => c.AddParagraph(
                            p => p.SetRun(
                                new RunBuilder()
                                .SetText($"Limit {i}")
                                .SetFontSize(9)
                                .SetFontFamily(FontsFamily.TimesNewRoman)
                                )
                            .SetAlignment(HorizontalAlignments.Center)
                            )
                        .SetWidth(
                            w =>
                            {
                                w.Width = "2";
                                w.Type = TypeWidthTable.Pct;
                            }
                            )
                    )
                    .AddCell(
                        c => c.AddParagraph(
                            p => p.SetRun(
                                new RunBuilder()
                                .SetText(i.ToString())
                                .SetFontSize(9)
                                .SetFontFamily(FontsFamily.TimesNewRoman)
                                .SetColor(Color.FromHex("#12ff5f"))
                                )
                            .SetAlignment(HorizontalAlignments.Center)
                            )

                    )
                );

        }


        table.Rows[0].Cells[1].Paragraphs[0].Runs[0].Color = Color.FromHex("#ff0000");
    }
}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}
