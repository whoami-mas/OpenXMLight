using OpenXMLight;
using OpenXMLight.Configurations.Elements.Charts;
using OpenXMLight.Spreadsheet;
using OpenXMLight.Spreadsheet.Elements;
using OpenXMLight.Spreadsheet.Formatting;

try
{
    using(var document = new WordDocument("test.docx", true))
    {
        List<ChartData> data = new List<ChartData>()
        {
            new ChartData()
            {
                Title = "Line1",
                Labels = new string[2] { "Element1", "Element2"},
                Data = new double[2] { 19, 13}
            },
            new ChartData()
            {
                Title = "Line2",
                Labels = new string[2] { "Element1", "Element2" },
                Data = new double[2] { 24, 10},
                TypeValueSeries = TypeSeries.General,
                orientationY = Orientation.Right
            }
        };
        
        ChartBuilder builder = new LineChart().SetTitle("Title chart").SetData(data).SetIsRightAxis(true, TypeValue.General);

        document.BuildChart(builder);

        document.Save();
    }
}
catch(Exception ex)
{
    Console.WriteLine(ex.Message);
}
