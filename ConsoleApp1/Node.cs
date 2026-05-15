using OpenXMLight;
using OpenXMLight.Configurations.Elements;
using OpenXMLight.Configurations.Elements.TableElements;
using OpenXMLight.Configurations.Elements.TableElements.Models;
using OpenXMLight.Configurations.Formatting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    public class Node
    {
        public void CreateNewDocument(WordDocument document)
        {

            document.SettingsDocument.Orientation = OrientationPage.Landscape;
            document.SettingsDocument.MarginTop = 55;


            Endnote endnote = document.AddEndnote("В(П) – вводный (первичный) инструктаж;" +
                    " Ц(В) – целевой (внеплановый) инструктаж;" +
                    " ПК – повышение квалификации (плановый) инструктаж");

            //Main table
            Table table = document.AddTable()
                .SetWidth(w =>
                {
                    w.Width = "100";
                    w.Type = TypeWidthTable.Pct;
                })
                .IsFixed(true)
                .SetBorders(
                        b => new Borders()
                        {
                            LineWidth = 1,
                            LineType = BordersType.Single
                        })
                .AddRows(
                r =>
                    r
                        .AddCell(
                            c =>
                                c.AddParagraph(
                                    p =>
                                        p.SetRun(
                                            new RunBuilder()
                                                .SetText("Дата")
                                                .SetBold(true)
                                                .SetFontSize(11)
                                                .SetFontFamily(FontsFamily.TimesNewRoman)
                                            )
                                        .SetAlignment(HorizontalAlignments.Center)
                                        .SetSpacingBetweenLines(new SpacingBetweenLines()
                                        {
                                            After = 0,
                                            Before = 0,
                                            Line = 250
                                        })
                                )
                                .SetVerticalAlignment(VerticalAlignments.Center)
                                .SetWidth(
                                    w =>
                                    {
                                        w.Width = "8,4";
                                        w.Type = TypeWidthTable.Pct;
                                    })
                        )
                        .AddCell(
                            c =>
                                c.AddParagraph(
                                    p =>
                                        p.SetRun(
                                            new RunBuilder()
                                                .SetText("Фамилия, имя, отчество обучаемого")
                                                .SetBold(true)
                                                .SetEndnote(endnote)
                                                .SetFontSize(11)
                                                .SetFontFamily(FontsFamily.TimesNewRoman)
                                            )
                                        .SetAlignment(HorizontalAlignments.Center)
                                        .SetSpacingBetweenLines(new SpacingBetweenLines()
                                        {
                                            After = 0,
                                            Before = 0,
                                            Line = 250
                                        })
                                )
                                .SetVerticalAlignment(VerticalAlignments.Center)
                                .SetWidth(
                                    w =>
                                    {
                                        w.Width = "20,8";
                                        w.Type = TypeWidthTable.Pct;
                                    })
                        )
                        .AddCell(
                            c =>
                                c.AddParagraph(
                                    p =>
                                        p.SetRun(
                                            new RunBuilder()
                                                .SetText("Год рождения")
                                                .SetBold(true)
                                                .SetFontSize(11)
                                                .SetFontFamily(FontsFamily.TimesNewRoman)
                                            )
                                        .SetAlignment(HorizontalAlignments.Center)
                                        .SetSpacingBetweenLines(new SpacingBetweenLines()
                                        {
                                            After = 0,
                                            Before = 0,
                                            Line = 250
                                        })
                                )
                                .SetVerticalAlignment(VerticalAlignments.Center)
                                .SetWidth(
                                    w =>
                                    {
                                        w.Width = "5,9";
                                        w.Type = TypeWidthTable.Pct;
                                    })
                        )
                )
                .AddRows(
                r =>
                    r
                        .AddCell(
                            c =>
                                c.AddParagraph(
                                    p =>
                                        p.SetRun(
                                            new RunBuilder()
                                                .SetText("")
                                                .SetBold(true)
                                                .SetFontSize(11)
                                                .SetFontFamily(FontsFamily.TimesNewRoman)
                                            )
                                        .SetSpacingBetweenLines(new SpacingBetweenLines()
                                        {
                                            After = 0,
                                            Before = 0,
                                            Line = 250
                                        })
                                )
                        )
                        .AddCell(
                            c =>
                                c.AddParagraph(
                                    p =>
                                        p.SetRun(
                                            new RunBuilder()
                                                .SetText("")
                                                .SetBold(true)
                                                .SetFontSize(11)
                                                .SetFontFamily(FontsFamily.TimesNewRoman)
                                            )
                                        .SetSpacingBetweenLines(new SpacingBetweenLines()
                                        {
                                            After = 0,
                                            Before = 0,
                                            Line = 250
                                        })
                                )
                        )
                        .AddCell(
                            c =>
                                c.AddParagraph(
                                    p =>
                                        p.SetRun(
                                            new RunBuilder()
                                                .SetText("")
                                                .SetBold(true)
                                                .SetFontSize(11)
                                                .SetFontFamily(FontsFamily.TimesNewRoman)
                                            )
                                        .SetSpacingBetweenLines(new SpacingBetweenLines()
                                        {
                                            After = 0,
                                            Before = 0,
                                            Line = 250
                                        })
                                )
                        )
                )
                .AddRows(
                r =>
                    r
                        .AddCell(
                            c =>
                                c.AddParagraph(
                                    p =>
                                        p.SetRun(
                                            new RunBuilder()
                                                .SetText("1")
                                                .SetBold(true)
                                                .SetFontSize(11)
                                                .SetFontFamily(FontsFamily.TimesNewRoman)
                                            )
                                        .SetAlignment(HorizontalAlignments.Center)
                                        .SetSpacingBetweenLines(new SpacingBetweenLines()
                                        {
                                            After = 0,
                                            Before = 0,
                                            Line = 250
                                        })
                                )
                        )
                        .AddCell(
                            c =>
                                c.AddParagraph(
                                    p =>
                                        p.SetRun(
                                            new RunBuilder()
                                                .SetText("2")
                                                .SetBold(true)
                                                .SetFontSize(11)
                                                .SetFontFamily(FontsFamily.TimesNewRoman)
                                            )
                                        .SetAlignment(HorizontalAlignments.Center)
                                        .SetSpacingBetweenLines(new SpacingBetweenLines()
                                        {
                                            After = 0,
                                            Before = 0,
                                            Line = 250
                                        })
                                )
                        )
                        .AddCell(
                            c =>
                                c.AddParagraph(
                                    p =>
                                        p.SetRun(
                                            new RunBuilder()
                                                .SetText("3")
                                                .SetBold(true)
                                                .SetFontSize(11)
                                                .SetFontFamily(FontsFamily.TimesNewRoman)
                                            )
                                        .SetAlignment(HorizontalAlignments.Center)
                                        .SetSpacingBetweenLines(new SpacingBetweenLines()
                                        {
                                            After = 0,
                                            Before = 0,
                                            Line = 250
                                        })
                                )
                        )
                )
                .SetMargin("0,15", "0,10", "0,15", "0");
        }
    }
}
