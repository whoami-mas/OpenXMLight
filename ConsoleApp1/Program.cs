using ConsoleApp1;
using OpenXMLight;
using OpenXMLight.Configurations;
using OpenXMLight.Configurations.Elements;
using OpenXMLight.Configurations.Elements.Table;
using OpenXMLight.Configurations.Formatting;

try
{
    //string path = @"C:\Users\bushk\Desktop\Reportings risks\act_rep_4_2025.xlsx";
    string path = @"testingTable.docx";

    using (WordDocument document = new WordDocument(path, true))
    {
        document.SettingsDocument.Orientation = OrientationPage.Landscape;
        document.SettingsDocument.MarginTop = 50;
        document.SettingsDocument.MarginBottom = 28;
        document.SettingsDocument.MarginFooter = 47;
        document.SettingsDocument.MarginGutter = 0;
        document.SettingsDocument.MarginHeader = 47;
        document.SettingsDocument.MarginLeft = 78;
        document.SettingsDocument.MarginRight = 78;

        int width = document.SettingsDocument.WidthPage;

        Row row1 = new Row();
        row1.Cells = new CellCollection(
            new Cell(new Text("1")),
            new Cell(new Text("2")),
            new Cell(new Text("3"))
            );
        Row row2 = new Row();
        row2.Cells = new CellCollection(
            new Cell(new Text("4")),
            new Cell(new Text("5")).Merge(1)
            );

        Table tbl = new TableBuilder().AppendRows(row1, row2).SetTableProperties(new TableProperties() { MarginCell = 17});
        document.AddTable(tbl);

    }
}
catch(Exception ex)
{
    Console.WriteLine(ex.Message);
}
