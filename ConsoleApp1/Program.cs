using ConsoleApp1;
using OpenXMLight;
using OpenXMLight.Spreadsheet.Elements;
using OpenXMLight.Spreadsheet.Formatting;


try
{
    //string path = @"C:\Users\bushk\Desktop\Reportings risks\act_rep_4_2025.xlsx";
    //string path = @"testingTable.docx";
    string path = @"C:\Users\bushk\Desktop\Reportings risks\for_test\2Org\МКК-Гудмэн-Отчет-мфд-3-кв.xlsx";
    string pathexc = @"C:\Users\bushk\Desktop\Reportings risks\for_test\2Org\Отчет-мфд-4-кв.xlsx";


    string pathF1 = "result1.txt";
    string pathF2 = "result2.txt";

    using (ExcelDocument document = new ExcelDocument(pathexc))
    {
        File.WriteAllText(pathF1, $"Start {DateTime.Now}");
        Sheet activeSheet = document.Sheets[0];

        for (int row_index = 1; row_index <= activeSheet.Rows.Count; row_index++)
        {
            List<string> row_values = new();

            for (int col_index = 1; col_index <= activeSheet.Rows[row_index].CountCell; col_index++)
            {
                string? value = activeSheet.Cells[row_index, col_index].Value?.ToString();

                if (string.IsNullOrWhiteSpace(value))
                    continue;
                else
                    row_values.Add(value);
            }

            if(row_values != null)
            {
                string values = string.Join(" | ", row_values);

                File.AppendAllText(pathF1, values+"\n");
            }
        }
    }

    using (ExcelDocument document = new ExcelDocument(path))
    {
        File.WriteAllText(pathF2, $"Start {DateTime.Now}");
        Sheet activeSheet = document.Sheets[0];

        for (int row_index = 1; row_index <= activeSheet.Rows.Count; row_index++)
        {
            List<string> row_values = new();

            for (int col_index = 1; col_index <= activeSheet.Rows[row_index].CountCell; col_index++)
            {
                string? value = activeSheet.Cells[row_index, col_index].Value?.ToString();

                if (string.IsNullOrWhiteSpace(value))
                    continue;
                else
                    row_values.Add(value);
            }

            if (row_values != null)
            {
                string values = string.Join(" | ", row_values);

                File.AppendAllText(pathF2, values + "\n");
            }
        }
    }
}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}
