using ConsoleApp1;
using OpenXMLight;
using OpenXMLight.Configurations.Elements;
using OpenXMLight.Configurations.Elements.TableElements.Models;
using OpenXMLight.Configurations.Formatting;


try
{ 
    string pathexc = @"F:\тестовые проекты\ConsoleApp1\тест\test.xlsx";

    using (ExcelDocument excel = new ExcelDocument(pathexc, true))
    {
        var activeSheet = excel.Sheets[0];

        activeSheet.Cells[1, 1].Value = DateTime.Now;
    }

}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}
