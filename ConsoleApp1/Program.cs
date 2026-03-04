using ConsoleApp1;
using OpenXMLight;
using OpenXMLight.Spreadsheet.Elements;
using OpenXMLight.Spreadsheet.Formatting;


try
{
    //string path = @"C:\Users\bushk\Desktop\Reportings risks\act_rep_4_2025.xlsx";
    //string path = @"testingTable.docx";
    string path = @"F:\тестовые проекты\ConsoleApp1\тест\testing.docx";
    string pathexc = @"F:\тестовые проекты\ConsoleApp1\тест\test.xlsx";


    using (ExcelDocument document = new ExcelDocument(pathexc, true))
    {
        Sheet activeSheet = document.Sheets[0];
        activeSheet.Name = "Паспорт";

        Cell cell = activeSheet.Cells[1, 1];
        cell.Value = $"ПАСПОРТ РИСКА - НАКОПЛЕННЫЕ РИСК-СОБЫТИЯ ЗА 2026 ГОД";
        activeSheet.Cells[1, 1, 1, 11].Merge();

        cell.Style.Font.Size = 16;
        cell.Style.Font.Bold = true;
        cell.Style.Font.CommitChange();
        cell.Style.Horizontal = HorizontalAlignments.Center;

        activeSheet.Cells[2, 1].Width = 13;
        activeSheet.Cells[2, 2].Width = 15;
        activeSheet.Cells[2, 3].Width = 23;
        activeSheet.Cells[2, 4].Width = 17.5;
        activeSheet.Cells[2, 5].Width = 33;
        activeSheet.Cells[2, 6].Width = 21;
        activeSheet.Cells[2, 7].Width = 55.5;
        activeSheet.Cells[2, 8].Width = 28;
        activeSheet.Cells[2, 9].Width = 28;
        activeSheet.Cells[2, 10].Width = 22.5;
        activeSheet.Cells[2, 11].Width = 17;

        activeSheet.Cells[2, 1].Value = "Порядковый номер";
        activeSheet.Cells[2, 2].Value = "Наименование риска";
        activeSheet.Cells[2, 3].Value = "Триггер (условие фиксации риск-события)";
        activeSheet.Cells[2, 4].Value = "Дата фиксации риск-события";
        activeSheet.Cells[2, 5].Value = "Описание риск-инцидента";
        activeSheet.Cells[2, 6].Value = "Оценка значимости (уровень ущерба)";
        activeSheet.Cells[2, 7].Value = "Мероприятия и/или процедуры по управлению риском, минимизация остаточного риска";
        activeSheet.Cells[2, 8].Value = "Владелец риска";
        activeSheet.Cells[2, 9].Value = "Перечень источников информации, используемых для идентификации и оценки риска";
        activeSheet.Cells[2, 10].Value = "Отметка о выполнении мероприятия";
        activeSheet.Cells[2, 11].Value = "Примечание срок исполнения мероприятий по управлению рисками";

        var headerTableRange = activeSheet.Cells[2, 1, 2, 11];
        headerTableRange.SetFont(
            f =>
            {
                f.Bold = true;
                f.Size = 11;
            }
        );

        for (int i = 3; i < 3 + 3; i++)
        {
            activeSheet.Cells[i, 1].Value = $"{i - 2}";
            activeSheet.Cells[i, 2].Value = $"Группа {i - 3}";
            activeSheet.Cells[i, 3].Value = $"Триггер {i - 3}";
            activeSheet.Cells[i, 4].Value = $"Фиксация даты {i - 3}";
            activeSheet.Cells[i, 5].Value = $"Описания {i - 3}";
            activeSheet.Cells[i, 6].Value = $"Оценка {i - 3}";
            activeSheet.Cells[i, 7].Value = $"Событие {i - 3}";
            activeSheet.Cells[i, 8].Value = $"Главный {i - 3}";
            activeSheet.Cells[i, 9].Value = $"Список {i - 3}";
            activeSheet.Cells[i, 10].Value = $"Вып {i - 3}";
            activeSheet.Cells[i, 11].Value = $"note {i - 3}";
        }

        var tableRange = activeSheet.Cells[2, 1, 3 + 2, 11];
        tableRange.SetBorder(
            b =>
            {
                b.Left.Type = BorderStyle.Single;
                b.Right.Type = BorderStyle.Single;
                b.Top.Type = BorderStyle.Single;
                b.Bottom.Type = BorderStyle.Single;
            }
        );
        tableRange.SetHorizontalAlignment(HorizontalAlignments.Center);
        tableRange.SetVerticalAlignment(VerticalAlignments.Center);
        tableRange.SetWrapText(true);
    }
}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}
