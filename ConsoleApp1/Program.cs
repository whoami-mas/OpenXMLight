using ConsoleApp1;
using OpenXMLight;
using OpenXMLight.Configurations.Formatting;
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
        activeSheet.Name = "Сотрудники";

        var range = activeSheet.Cells["A1:C3"];
        range.SetBorder(f =>
        {
            f.Left.Type = BorderStyle.Single;
            f.Left.Color = Color.FromHex("#fff");

            f.Right.Type = BorderStyle.Single;
            f.Right.Color = Color.FromHex("#fff");

            f.Top.Type = BorderStyle.Single;
            f.Top.Color = Color.FromHex("#fff");

            f.Bottom.Type = BorderStyle.Single;
            f.Bottom.Color = Color.FromHex("#fff");
        });

        range[1, 1].Value = "Hello World!";
    }
}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}
