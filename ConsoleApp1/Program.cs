using ConsoleApp1;
using OpenXMLight;
using OpenXMLight.Spreadsheet.Elements;
using OpenXMLight.Spreadsheet.Formatting;


try
{ 
    string pathexc = @"F:\тестовые проекты\ConsoleApp1\тест\test.xlsx";

    using (ExcelDocument excel = new ExcelDocument(pathexc, true))
    {
        Sheet activeSheet = excel.Sheets[0];
        activeSheet.Name = "Лимиты";

        Cell titleCell = activeSheet.Cells[2, 2];
        titleCell.Value = $"Проект приказа по утверждению квартальных лимитов риска на 2026 год";
        activeSheet.Cells[2, 2, 2, 8].Merge();

        titleCell.Style.Font.Bold = true;
        titleCell.Style.Font.Size = 16;
        titleCell.Style.Font.CommitChange();
        titleCell.Style.Horizontal = HorizontalAlignments.Center;

        activeSheet.Cells[3, 3].Style.TextRotation = 90;
        activeSheet.Cells[4, 2].Width = 8.43;
        activeSheet.Cells[4, 3].Width = 20.14;
        activeSheet.Cells[4, 4].Width = 6.57;
        activeSheet.Cells[4, 5].Width = 43.43;
        activeSheet.Cells[4, 6].Width = 19.43;
        activeSheet.Cells[4, 7].Width = 40.14;

        activeSheet.Cells[4, 2].Value = "Риск";
        activeSheet.Cells[4, 3].Value = "Лимит на квартал";
        activeSheet.Cells[4, 4].Value = "№";
        activeSheet.Cells[4, 5].Value = "Ключевой индикатор риска";
        activeSheet.Cells[4, 6].Value = "Владелец риска";
        activeSheet.Cells[4, 7].Value = "Мероприятия по управлению риском (стандартные)";

        Cells headerTable = activeSheet.Cells[4, 2, 4, 7];
        headerTable.SetBorder(
            b =>
            {
                b.Left.Type = BorderStyle.Single;
                b.Top.Type = BorderStyle.Single;
                b.Right.Type = BorderStyle.Single;
                b.Bottom.Type = BorderStyle.Single;
            });
        headerTable.SetFont(
            f =>
            {
                f.Bold = true;
                f.Size = 11;
            });
        headerTable.SetWrapText(true);
        headerTable.SetHorizontalAlignment(HorizontalAlignments.Center);
        headerTable.SetVerticalAlignment(VerticalAlignments.Center);

        int number = 1;
        for (int i = 5; i < 16; i++)
        {
            activeSheet.Cells[i, 2].Value = "name group risk";
            activeSheet.Cells[i, 3].Value = i;
            activeSheet.Cells[i, 4].Value = number;
            activeSheet.Cells[i, 5].Value = $"risk name {i}";
            activeSheet.Cells[i, 6].Value = "Риск-менеджер";

            activeSheet.Cells[i, 2].Style.TextRotation = 90;

            number++;
        }

        Cells tableRange = activeSheet.Cells[5, 2, number + 3, 7];
        tableRange.SetBorder(
            b =>
            {
                b.Left.Type = BorderStyle.Single;
                b.Top.Type = BorderStyle.Single;
                b.Right.Type = BorderStyle.Single;
                b.Bottom.Type = BorderStyle.Single;
            });
        tableRange.SetVerticalAlignment(VerticalAlignments.Center);
        tableRange.SetHorizontalAlignment(HorizontalAlignments.Center);

        activeSheet.Cells[17, 3].Value = "УТВЕРЖДАЮ";
        activeSheet.Cells[17, 5].Value = "Руководитель МКК";
        activeSheet.Cells[17, 6].Value = "___________________";
    }

}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}
