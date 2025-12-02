using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXmlDrawing = DocumentFormat.OpenXml.Drawing;
using OpenXmlF = DocumentFormat.OpenXml;
using OpenXmlChart = DocumentFormat.OpenXml.Drawing.Charts;
using System.Globalization;

namespace OpenXMLight.Configurations.Elements.Charts
{
    public class Chart
    {
        private string title;
        private List<ChartData> data;
        private int width;
        private int height;

        
        public string Title { get => title; set => title = value; }
        public List<ChartData> Data { get => data; set => data = value; }
        public int Width { get =>  width;
            set
            {
                width = value;

                WidthLong = (long)(width * 15) * 1000;
            }
        }
        public int Height { get => height;
            set
            {
                height = value;

                HeightLong = (long)(height * 15) * 1000;
            }
        }


        internal OpenXmlChart.ChartSpace ChartSpaceXml { get; set; }
        internal OpenXmlChart.Chart ChartXml { get; set; }
        internal OpenXmlF.Int64Value WidthLong { get; private set; }
        internal OpenXmlF.Int64Value HeightLong { get; private set; }

        internal Chart(OpenXmlChart.ChartSpace? chartSpaceXml = default)
        {
            this.Width = 367;
            this.Height = 267;

            this.ChartSpaceXml = chartSpaceXml ??= new OpenXmlChart.ChartSpace();
            this.ChartXml = this.ChartSpaceXml.Elements<OpenXmlChart.Chart>().FirstOrDefault() ?? this.ChartSpaceXml.AppendChild(new OpenXmlChart.Chart());
        }
    }
}
