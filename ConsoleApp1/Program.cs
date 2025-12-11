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
        //document.SettingsDocument.Orientation = OrientationPage.Landscape;

        //Row row1 = new Row();
        //row1.Cells = new CellCollection(
        //    new Cell(new Text("", textProp: new TextProperties()
        //    {
        //        FontSize = 16,
        //        Bold = true,
        //        FontFamily = FontsFamily.TimesNewRoman,
        //        HAlignment = HorizonatalAlignments.Center,
        //        SpBetLines = new SpacingBetweenLines()
        //        { After = 0, Before = 0 }
        //    })),
        //    new Cell(new Text("Начат:", textProp: new TextProperties()
        //    {
        //        FontSize = 16,
        //        Bold = true,
        //        FontFamily = FontsFamily.TimesNewRoman,
        //        HAlignment = HorizonatalAlignments.Center,
        //        SpBetLines = new SpacingBetweenLines()
        //        { After = 0, Before = 0 }
        //    }), mergeColumn: 2),
        //    new Cell(new Text("hlkhkjhlkhlkhjk"))
        //    );
        //Row row2 = new Row();
        //row2.Cells = new CellCollection(
        //    new Cell(new Text("", textProp: new TextProperties()
        //    {
        //        FontSize = 16,
        //        Bold = true,
        //        FontFamily = FontsFamily.TimesNewRoman,
        //        HAlignment = HorizonatalAlignments.Center,
        //        SpBetLines = new SpacingBetweenLines()
        //        { After = 0, Before = 0 }
        //    })),
        //    new Cell(new Text("Окончен:", textProp: new TextProperties()
        //    {
        //        FontSize = 16,
        //        Bold = true,
        //        FontFamily = FontsFamily.TimesNewRoman,
        //        HAlignment = HorizonatalAlignments.Center,
        //        SpBetLines = new SpacingBetweenLines()
        //        { After = 0, Before = 0 }
        //    })),
        //    new Cell(new Text($"__.__.____", textProp: new TextProperties()
        //    {
        //        FontSize = 16,
        //        FontFamily = FontsFamily.TimesNewRoman,
        //        HAlignment = HorizonatalAlignments.Center,
        //        SpBetLines = new SpacingBetweenLines()
        //        { After = 0, Before = 0 }
        //    }))
        //    );

        //Table table_date = new TableBuilder().AppendRows(row1, row2)
        //    .SetTableGrid(100, 200, 200);

        //document.AddTable(table_date);
        document.SettingsDocument.Orientation = OrientationPage.Landscape;
        document.SettingsDocument.MarginTop = 425;
        document.SettingsDocument.MarginLeft = 1170;
        document.SettingsDocument.MarginRight = 1170;

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

        Table tbl = new TableBuilder().AppendRows(row1, row2);
        document.AddTable(tbl);

    }
}
catch(Exception ex)
{
    Console.WriteLine(ex.Message);
}
