using ConsoleApp1;
using OpenXMLight;
using OpenXMLight.Configurations;
using OpenXMLight.Configurations.Elements;
using OpenXMLight.Configurations.Elements.Table;
using OpenXMLight.Configurations.Formatting;

try
{
    //string path = @"C:\Users\bushk\Desktop\Reportings risks\act_rep_4_2025.xlsx";
    //string path = @"testingTable.docx";
    string path = @"F:\тестовые проекты\NewServiceRisksPlus\NewServiceRisksPlus\bin\Debug\net8.0-windows7.0\archives\a_44\2025\Журнал проведения обучений 2025.docx";

    using (WordDocument document = new WordDocument(path))
    {
        RowCollection rows = document.Tables[1].Rows;

        rows[7].Cells[1].Text.Content = "hello";
    }
}
catch(Exception ex)
{
    Console.WriteLine(ex.Message);
}
