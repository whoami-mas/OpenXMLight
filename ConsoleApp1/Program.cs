using ConsoleApp1;
using DocumentFormat.OpenXml;
using OpenXMLight;
using OpenXMLight.Configurations.Elements;
using OpenXMLight.Configurations.Elements.TableElements;
using OpenXMLight.Configurations.Elements.TableElements.Models;
using OpenXMLight.Configurations.Formatting;
using System.Reflection;


string path = @"F:\тестовые проекты\ConsoleApp1\тест\template.docx";
FileInfo file = new FileInfo(path);

try
{
    using (var document = new WordDocument(path, false))
    {
        if (true)
            new Node().CreateNewDocument(document);

        Table table = document.Tables[0];

    }

}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}
