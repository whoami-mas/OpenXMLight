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

        TextProperties txtProp = new TextProperties()
        {
            FontSize = 16,
            FontFamily = FontsFamily.TimesNewRoman,
            Bold = true,
            HAlignment = HorizonatalAlignments.Center,
        };
        TextProperties txtProp1 = new TextProperties()
        {
            FontSize = 24,
            FontFamily = FontsFamily.TimesNewRoman,
            Bold = true,
            HAlignment = HorizonatalAlignments.Center,
        };
        TextProperties txtProp2 = new TextProperties()
        {
            FontSize = 16,
            FontFamily = FontsFamily.TimesNewRoman,
            HAlignment = HorizonatalAlignments.Center,
        };

        document.AddText(new Text($"ООО МКК «test»", textProp: txtProp));

        txtProp1.SpBetLines.Before = 220;
        document.AddText(new Text("ЖУРНАЛ", textProp: txtProp1));

        txtProp.SpBetLines.After = 150;
        document.AddText(new Text("Журнал учета проведения обучений по Риск-менеджменту", textProp: txtProp));

        Table table_date = new Table();
        Row row1 = new Row();
        row1.Cells = new CellCollection(
            new Cell(new Text("")),
            new Cell(new Text("Начат:", textProp: txtProp)),
            new Cell(new Text($"{DateTime.Now.ToString("dd.MM.yyyy")}", textProp: txtProp2))
            );
        table_date.Rows.Add(row1);
        Row row2 = new Row();
        row2.Cells = new CellCollection(
            new Cell(new Text("")),
            new Cell(new Text("Окончен:", textProp: txtProp)),
            new Cell(new Text($"__.__.____", textProp: txtProp2))
            );
        table_date.Rows.Add(row2);

        document.AddTable(table_date);
    }
}
catch(Exception ex)
{
    Console.WriteLine(ex.Message);
}
