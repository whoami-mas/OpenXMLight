using ConsoleApp1;
using OpenXMLight;
using OpenXMLight.Configurations;
using OpenXMLight.Configurations.Elements;
using OpenXMLight.Configurations.Elements.TableElements.Formattings.MarginComponents;
using OpenXMLight.Configurations.Elements.TableElements.Models;
using OpenXMLight.Configurations.Formatting;

try
{
    //string path = @"C:\Users\bushk\Desktop\Reportings risks\act_rep_4_2025.xlsx";
    //string path = @"testingTable.docx";
    string path = @"F:\тестовые проекты\NewServiceRisksPlus\NewServiceRisksPlus\bin\Debug\net8.0-windows7.0\archives\a_44\2025\Журнал проведения обучений 2025.docx";

    using (WordDocument document = new WordDocument(path, true))
    {
        #region template
        //TableTest tbl = document.AddTable()
        //    .SetWidth(w =>
        //    {
        //        w.Width = "15";
        //        w.Type = TypeWidthTable.Cm;
        //    })
        //    .SetBorders(b =>
        //    {
        //        b.LineWidth = 0.5;
        //        b.LineType = BordersType.Single;
        //    }).
        //    SetMargin("0,19", "0", "0,19", "0")
        //    .AddRows(
        //    row =>
        //    row.AddCell(
        //        cell => cell.AddParagraph(
        //            p => p.SetRun(
        //                new RunBuilder().SetText("1")
        //                )
        //            )
        //        )
        //    .AddCell(cell => cell.AddParagraph(
        //            p => p.SetRun(
        //                new RunBuilder().SetText("2")
        //                )
        //            )
        //        )
        //    .AddCell(cell => cell.AddParagraph(
        //            p => p.SetRun(
        //                new RunBuilder().SetText("3")
        //                )
        //            )
        //        )
        //    )
        //    .AddRows(
        //    row => row.AddCell(
        //        cell => cell.AddParagraph(
        //            p => p.SetRun(
        //                new RunBuilder().SetText("4")
        //                )
        //            )
        //        ).AddCell(cell => cell.AddParagraph(
        //            p => p.SetRun(
        //                new RunBuilder().SetText("5")
        //                )
        //            )
        //        ).AddCell(cell => cell.AddParagraph(
        //            p => p.SetRun(
        //                new RunBuilder().SetText("6")
        //                )
        //            )
        //        )
        //    )
        //    .AddRows(row => row.AddCell(
        //        cell => cell.AddParagraph(
        //            p => p.SetRun(
        //                new RunBuilder().SetText("7")
        //                )
        //            )
        //        ).AddCell(cell => cell.AddParagraph(
        //            p => p.SetRun(
        //                new RunBuilder().SetText("8")
        //                )
        //            )
        //        ).AddCell(cell => cell.AddParagraph(
        //            p => p.SetRun(
        //                new RunBuilder().SetText("9")
        //                )
        //            )
        //        )
        //    )
        //    .Merge(0, 1, 0, 2)
        //    .IsFixed(true);
        #endregion

        EndnoteTest endnote = document.AddEndnote("Hello World!");

        Paragraph p = document.AddParagraph()
            .SetRun(
                new RunBuilder().SetText("Testing endnote"),
                new RunBuilder().SetEndnote(endnote)
            );
    }
}
catch(Exception ex)
{
    Console.WriteLine(ex.Message);
}
