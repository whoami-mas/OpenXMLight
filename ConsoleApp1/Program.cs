using ConsoleApp1;
using OpenXMLight;
using OpenXMLight.Spreadsheet.Elements;

try
{
    //string path = @"C:\Users\bushk\Desktop\Reportings risks\act_rep_4_2025.xlsx";
    //string path = @"testingTable.docx";
    string path = @"F:\тестовые проекты\ConsoleApp1\тест\testing.docx";
    string pathexc = @"F:\тестовые проекты\ConsoleApp1\тест\test.xlsx";
    //string pathexc = @"F:\тестовые проекты\ConsoleApp1\тест\template_import.xlsx";


    using (ExcelDocument document = new ExcelDocument(pathexc, true))
    {
        Sheet activeSheet = document.Sheets[0];
        activeSheet.Name = "Сотрудники";

        activeSheet.Cells[1, 1].Value = "ФИО сотрудника";
        activeSheet.Cells[1, 2].Value = "Должность сотрудника";
        activeSheet.Cells[1, 3].Value = "Дата рождения сотрудника";
        activeSheet.Cells[1, 4].Value = "Дата приема сотрудника на работу";

        activeSheet.Cells[2, 1].Value = "Иванов Иван Иванович";
        activeSheet.Cells[2, 2].Value = "Фронтенщик";
        activeSheet.Cells[2, 3].Value = DateTime.Parse("12.12.1999");
        activeSheet.Cells[2, 4].Value = DateTime.Now;

        activeSheet.Cells.AutoFitColumns();

        //Console.WriteLine(activeSheet.Cells[2, 4].Value);
    }
}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}
