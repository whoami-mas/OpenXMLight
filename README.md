# OpenXMLight
<h3>Library for easier work with XML Office</h3>
Format support .docx, .xlsx

## Word
### Example of creating a graph 📈
```C#
WordDocument document = new WordDocument("example.docx");
ChartBuilder builder = new LineChart().SetTitle("Title chart").SetData(data);
document.BuildChart(builder);
document.Save();
```
### Example of create table
```C#
WordDocument document = new WordDocument("example.docx");
Table tbl = document.AddTable()
    .AddRows(row =>
    row.AddCell(
        cell => cell.AddParagraph(
            p => p.SetRun(
                new RunBuilder().SetText("1")
                )
            )
        )
    .AddCell(cell => cell.AddParagraph(
            p => p.SetRun(
                new RunBuilder().SetText("2")
                )
            )
        )
    )
    .AddRows(
    row => row.AddCell(
        cell => cell.AddParagraph(
            p => p.SetRun(
                new RunBuilder().SetText("4")
                )
            )
        ).AddCell(cell => cell.AddParagraph(
            p => p.SetRun(
                new RunBuilder().SetText("5")
                )
            )
        )
    )
    .AddRows(row => row.AddCell(
        cell => cell.AddParagraph(
            p => p.SetRun(
                new RunBuilder().SetText("7")
                )
            )
        ).AddCell(cell => cell.AddParagraph(
            p => p.SetRun(
                new RunBuilder().SetText("8")
                )
            )
        )
    );
document.Save();
```

### Example added endnote
```C#
WordDocument document = new WordDocument("example.docx");
EndnoteTest endnote = document.AddEndnote("Hello World!");
Paragraph p = document.AddParagraph()
    .SetRun(
        new RunBuilder().SetText("Testing endnote"),
        new RunBuilder().SetEndnote(endnote)
    );
document.Save();
```
##Excel
