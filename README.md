# OpenXMLight
<h3>Library for easier work with XML Office</h3>
Format support .docx, .xlsx

<h3>Example of creating a graph</h3>

ChartBuilder builder = new LineChart().SetTitle("Title chart").SetData(data);
document.BuildChart(builder);
