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

        Row header = new Row();
        header.Cells = new CellCollection(
            new Cell(new Text("Дата", textProp: new TextProperties()
            {
                FontSize = 11,
                Bold = true,
                FontFamily = FontsFamily.TimesNewRoman,
                HAlignment = HorizonatalAlignments.Center
            }), vMerge: VerticalMerge.Start, vAlignment: VerticalAlignment.Center),
            new Cell(new Text("Фамилия, имя, отчество обучаемого", textProp: new TextProperties()
            {
                FontSize = 11,
                Bold = true,
                FontFamily = FontsFamily.TimesNewRoman,
                HAlignment = HorizonatalAlignments.Center
            }), vMerge: VerticalMerge.Start),
            new Cell(new Text("Год рождения", textProp: new TextProperties()
            {
                FontSize = 11,
                Bold = true,
                FontFamily = FontsFamily.TimesNewRoman,
                HAlignment = HorizonatalAlignments.Center
            }), vMerge: VerticalMerge.Start),
            new Cell(new Text("Должность обучаемого", textProp: new TextProperties()
            {
                FontSize = 11,
                Bold = true,
                FontFamily = FontsFamily.TimesNewRoman,
                HAlignment = HorizonatalAlignments.Center
            }), vMerge: VerticalMerge.Start),
            new Cell(new Text("Вид обучения", textProp: new TextProperties()
            {
                FontSize = 11,
                Bold = true,
                FontFamily = FontsFamily.TimesNewRoman,
                HAlignment = HorizonatalAlignments.Center
            }), vMerge: VerticalMerge.Start),
            new Cell(new Text("Основание проведения внепланового обучения", textProp: new TextProperties()
            {
                FontSize = 11,
                Bold = true,
                FontFamily = FontsFamily.TimesNewRoman,
                HAlignment = HorizonatalAlignments.Center
            }), vMerge: VerticalMerge.Start),
            new Cell(new Text("Фамилия, инициалы, должность обучающего", textProp: new TextProperties()
            {
                FontSize = 11,
                Bold = true,
                FontFamily = FontsFamily.TimesNewRoman,
                HAlignment = HorizonatalAlignments.Center
            }), vMerge: VerticalMerge.Start),
            new Cell(new Text("Подпись", textProp: new TextProperties()
            {
                FontSize = 11,
                Bold = true,
                FontFamily = FontsFamily.TimesNewRoman,
                HAlignment = HorizonatalAlignments.Center
            })).Merge(1)
        );
        Row header2 = new Row();
        header2.Cells = new CellCollection(
            new Cell(new Text(""), vMerge: VerticalMerge.Continue),
            new Cell(new Text(""), vMerge: VerticalMerge.Continue),
            new Cell(new Text(""), vMerge: VerticalMerge.Continue),
            new Cell(new Text(""), vMerge: VerticalMerge.Continue),
            new Cell(new Text(""), vMerge: VerticalMerge.Continue),
            new Cell(new Text(""), vMerge: VerticalMerge.Continue),
            new Cell(new Text(""), vMerge: VerticalMerge.Continue),
            new Cell(new Text("Обучающего", textProp: new TextProperties()
            {
                FontSize = 11,
                Bold = true,
                FontFamily = FontsFamily.TimesNewRoman,
                HAlignment = HorizonatalAlignments.Center
            })),
            new Cell(new Text("Обучаемого", textProp: new TextProperties()
            {
                FontSize = 11,
                Bold = true,
                FontFamily = FontsFamily.TimesNewRoman,
                HAlignment = HorizonatalAlignments.Center
            }))
        );

        Table table_data = new TableBuilder().AppendRows(header, header2)
            .SetTableGrid(83, 206, 58, 113, 75, 121, 132, 103, 96)
            .SetTableProperties(new TableProperties() { MarginCell = 5, Fixed = true });

        document.AddTable(table_data);

    }
}
catch(Exception ex)
{
    Console.WriteLine(ex.Message);
}
