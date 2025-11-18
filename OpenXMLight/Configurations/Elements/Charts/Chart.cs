using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXMLDrawing = DocumentFormat.OpenXml.Drawing;
using OpenXmlChart = DocumentFormat.OpenXml.Drawing.Charts;
using System.Globalization;

namespace OpenXMLight.Configurations.Elements.Charts
{
    public class Chart
    {
        private string title;
        private List<ChartData> data;

        public string Title { get => title; set => title = value; }
        public List<ChartData> Data { get => data; set => data = value; }

        internal OpenXmlChart.ChartSpace ChartSpaceXml { get; set; }
        internal OpenXmlChart.Chart ChartXml { get; set; }

        internal Chart(OpenXmlChart.ChartSpace? chartSpaceXml = default)
        {
            this.ChartSpaceXml = chartSpaceXml ??= new OpenXmlChart.ChartSpace();
            this.ChartXml = this.ChartSpaceXml.Elements<OpenXmlChart.Chart>().FirstOrDefault() ?? new OpenXmlChart.Chart();
        }
    }
}
