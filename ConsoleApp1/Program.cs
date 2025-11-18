using OpenXMLight;
using OpenXMLight.Configurations;
using OpenXMLight.Configurations.Formatting;
using OpenXMLight.Configurations.Elements.Table;
using OpenXMLight.Configurations.Elements;
using OpenXMLight.Configurations.Elements.Charts;

try
{
    using (var test = new WordDocument("testingWord.docx", true))
    {
        List<ChartData> chData = new()
        {
            new ChartData()
            {
                Title = "Величина портфеля займов, тыс. рублей",
                Labels = new string[3]{ "1 кв 2025", "2 кв 2025", "3 кв 2025"},
                Data = new double[3] {3.2, 1.2, 2 }
            },
            new ChartData()
            {
                Title = "Проблемный портфель NPL90+, тыс. рублей",
                Labels = new string[3]{ "1 кв 2025", "2 кв 2025", "3 кв 2025"},
                Data = new double[3] {1.2, 2.3, 5 }
            },
            new ChartData()
            {
                Title = "Удельный вес просроченных займов NPL90+, %",
                Labels = new string[3]{ "1 кв 2025", "2 кв 2025", "3 кв 2025"},
                Data = new double[3] {1.7, 3.6, 4 }
            }
        };

        ColumnChart chart = new ColumnChart("Создание тестового графика", chData);

        test.AddChart(chart);
    }
}
catch(Exception ex)
{
    Console.WriteLine(ex.Message);
}
