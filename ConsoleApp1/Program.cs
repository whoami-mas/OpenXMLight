using ConsoleApp1;
using OpenXMLight;
using OpenXMLight.Configurations.Elements.Table;
using OpenXMLight.Configurations.Formatting;
using OpenXMLight.Spreadsheet;
using OpenXMLight.Spreadsheet.Elements;
using OpenXMLight.Spreadsheet.Formatting;

try
{
    //string path = @"C:\Users\bushk\Desktop\Reportings risks\act_rep_4_2025.xlsx";
    string path = @"testingTable.docx";

    using (WordDocument doc = new WordDocument(path))
    {
        Table countTable = doc.Tables[0];

        string s = countTable.Rows[1].Cells[1].Text.Content;
    }
}
catch(Exception ex)
{
    Console.WriteLine(ex.Message);
}
