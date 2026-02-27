using ConsoleApp1;
using OpenXMLight;
using OpenXMLight.Spreadsheet.Elements;

try
{
    //string path = @"C:\Users\bushk\Desktop\Reportings risks\act_rep_4_2025.xlsx";
    //string path = @"testingTable.docx";
    string path = @"F:\тестовые проекты\ConsoleApp1\тест\testing.docx";
    //string pathexc = @"F:\тестовые проекты\ConsoleApp1\тест\test.xlsx";
    string pathexc = @"F:\тестовые проекты\ConsoleApp1\тест\template_import.xlsx";


    using (ExcelDocument document = new ExcelDocument(pathexc, false))
    {
        Sheet activeSheet = document.Sheets[0];

        activeSheet.Cells[5, 5].Value = DateTime.Now;
    }
}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}
