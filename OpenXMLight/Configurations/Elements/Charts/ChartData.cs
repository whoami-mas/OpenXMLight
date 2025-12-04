using OpenXMLight.Spreadsheet.Formatting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Configurations.Elements.Charts
{
    public enum Orientation
    {
        Left, Right
    }
    public enum TypeSeries
    {
        General, Percent
    }
    public class ChartData
    {
        public string Title { get; set; }
        public string[] Labels { get; set; }
        public double[] Data { get; set; }
        public Orientation orientationY { get; set; } = Orientation.Left;
        public TypeSeries TypeValueSeries { get; set; } = TypeSeries.General;
    }
}
