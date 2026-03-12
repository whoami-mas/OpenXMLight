using ConsoleApp1;
using OpenXMLight;
using OpenXMLight.Configurations.Elements;


try
{ 
    string pathexc = @"F:\тестовые проекты\ConsoleApp1\тест\test.docx";

    using (WordDocument word = new WordDocument(pathexc, true))
    {
        Endnote endnote = word.AddEndnote("testing");

        word.AddParagraph().SetRun(new RunBuilder().SetText("Hello World!").SetEndnote(endnote));
    }

}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}
