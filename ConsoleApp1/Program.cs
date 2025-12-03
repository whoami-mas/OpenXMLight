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
                Labels = new string[1] { "Element1"},
                Data = new double[1] { 9881.382}
            },
            new ChartData()
            {
                Title = "Line2",
                Labels = new string[1] { "Element2" },
                Data = new double[1] { 50.44},
            }
        };
        
        ChartBuilder builder = new PieChart().SetTitle("Title chart").SetData(data);

        document.BuildChart(builder);

        document.Save();
    }
}
catch(Exception ex)
{
    Console.WriteLine(ex.Message);
}
