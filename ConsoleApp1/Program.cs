using ConsoleApp1;
using OpenXMLight;
using OpenXMLight.Configurations.Elements.Charts;
using OpenXMLight.Spreadsheet;
using OpenXMLight.Spreadsheet.Elements;
using OpenXMLight.Spreadsheet.Formatting;

try
{
    string path = @"C:\Users\bushk\Desktop\Reportings risks\act_rep_4_2025.xlsx";

    List<Node> tree = new();
    using(var document = new ExcelDocument(path))
    {
        Sheet activeSheet = document.Sheets[0];
        
        int index = 1;
        for (int i = index; i <= activeSheet.Rows.Count; i++)
        {
            if (activeSheet.Cells[i, 1].Value != null && activeSheet.Cells[i, 1].Value.ToString().ToLower().Trim().StartsWith("раздел"))
            {
                Node chapter = new Node()
                {
                    indexRow = i,
                    name = activeSheet.Cells[i, 1].Value.ToString(),
                    nodes = new()
                };

                for (int j = chapter.indexRow + 1; j <= activeSheet.Rows.Count; j++)
                {
                    List<string> valueRow = new();

                    for(int col = 1; col <= activeSheet.Rows[j].CountCell; col++)
                    {
                        if (activeSheet.Cells[j, col].Value != null)
                            valueRow.Add(activeSheet.Cells[j, col].Value.ToString());
                    }

                    Node node = new Node() {indexRow = j };
                    if(valueRow.Count >= 3)
                    {
                        node.pointName = valueRow[0].Trim();
                        node.name = valueRow[1].Trim();
                        node.value = valueRow[2].Trim();
                    }

                    chapter.nodes.Add(node);
                    if (activeSheet.Cells[j, 1].Value != null && activeSheet.Cells[j, 1].Value.ToString().ToLower().Trim().StartsWith("раздел"))
                    {
                        index = j;
                        break;
                    }
                }

                tree.Add(chapter);
            }
        }

        foreach(Node node in tree)
        {
            File.AppendAllText("test.txt", string.Format("\t{0}\n", node.name));
            foreach (Node child in node.nodes)
                File.AppendAllText("test.txt", string.Format("\t\t-{0} : {1}\n", child.pointName, child.value));
        }
    }
}
catch(Exception ex)
{
    Console.WriteLine(ex.Message);
}
