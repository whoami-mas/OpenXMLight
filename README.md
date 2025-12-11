# OpenXMLight
<h3>Library for easier work with XML Office</h3>
Format support .docx, .xlsx

<h2>Word</h2>
<h3>Example of creating a graph 📈</h3>
<p>WordDocument document = new WordDocument("example.docx");</p>
<p>ChartBuilder builder = new LineChart().SetTitle("Title chart").SetData(data);</p>
<p>document.BuildChart(builder);</p>
<p>document.Save();</p>

### Example of create table 📈

```C#
Row row1 = new Row();
row1.Cells = new CellCollection(
    new Cell(new Text("1")),
    new Cell(new Text("2")),
    new Cell(new Text("3"))
    );
Row row2 = new Row();
row2.Cells = new CellCollection(
    new Cell(new Text("4")),
    new Cell(new Text("5")).Merge(1)
    );

Table tbl = new TableBuilder().AppendRows(row1, row2);
document.AddTable(tbl);
```
<h2>Excel</h2>
