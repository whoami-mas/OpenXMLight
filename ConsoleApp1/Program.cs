using OpenXMLight;
using OpenXMLight.Configurations.Elements.Charts;
using OpenXMLight.Spreadsheet;
using OpenXMLight.Spreadsheet.Elements;

try
{
    using(var document = new WordDocument("test.docx", true))
    {
        List<ChartData> data = new List<ChartData>()
        {
            new ChartData()
            {
                Title = "Line1",
                Labels = new string[2] { "Element1", "Element2" },
                Data = new double[2] { 9881.382, 10953.682}
            },
            new ChartData()
            {
                Title = "Line2",
                Labels = new string[2] { "Element1", "Element2" },
                Data = new double[2] { 50.44, 39.25},
                orientationY = Orientation.Right,
            },
            new ChartData()
            {
                Title = "Line3",
                Labels = new string[2] { "Element1", "Element2" },
                Data = new double[2] { 4983.71, 4299.85 }
            }
        };
        
        ChartBuilder builder = new LineChart().SetTitle("Title chart").SetData(data);

        document.BuildChart(builder);

        document.Save();
    }
}
catch(Exception ex)
{
    Console.WriteLine(ex.Message);
}
