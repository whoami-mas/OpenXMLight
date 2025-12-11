# OpenXMLight
<h3>Library for easier work with XML Office</h3>
Format support .docx, .xlsx

<h2>Word</h2>
<h3>Example of creating a graph 📈</h3>
<p>WordDocument document = new WordDocument("example.docx");</p>
<p>ChartBuilder builder = new LineChart().SetTitle("Title chart").SetData(data);</p>
<p>document.BuildChart(builder);</p>
<p>document.Save();</p>

<h3>Example of create table 📈</h3>
```python
def add(a, b):
    return a + b

print(add(2, 3))
<h2>Excel</h2>
