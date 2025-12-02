using OpenXMLight;
using OpenXMLight.Configurations.Elements.Charts;
using OpenXMLight.Spreadsheet;
using OpenXMLight.Spreadsheet.Elements;

try
{
    using(var document = new WordDocument("test.docx", true))
    {
        List<ChartData> d = new List<ChartData>()
        {
            new ChartData()
            {
                Title = "Линия1",
                Labels = new string[2] { "Элемент1", "Элемент2" },
                Data = new double[2] { 9881.382, 10953.682}
            },
            new ChartData()
            {
                Title = "Линия2",
                Labels = new string[2] { "Элемент1", "Элемент2" },
                Data = new double[2] { 50.44, 39.25},
                orientationY = Orientation.Right,
            },
            new ChartData()
            {
                Title = "Линия3",
                Labels = new string[2] { "Элемент1", "Элемент2" },
                Data = new double[2] { 4983.71, 4299.85 }
            }
        };
        
        ChartBuilder chartB = new LineChart().SetTitle("Тест").SetData(d);

        document.BuildChart(chartB);
    }
}
catch(Exception ex)
{
    Console.WriteLine(ex.Message);
}
