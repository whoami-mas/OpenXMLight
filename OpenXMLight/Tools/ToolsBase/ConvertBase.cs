using OpenXMLight.Configurations;
using OpenXMLight.Configurations.Formatting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Tools.ToolsBase
{
    public abstract class ConvertBase
    {


        public virtual T GetWidthOfType<T>(string width, TypeWidthTable type) where T : IConvertible
        {
            double result = double.Parse(width);

            if (type == TypeWidthTable.Pct)
                result *= 50;
            else if(type == TypeWidthTable.Cm)
                result *= Configuration.WidthCm;
            else
            {

            }

            return (T)Convert.ChangeType((short)result, typeof(T));
        }

        public virtual string ConvertWidth(string width, TypeWidthTable type)
        {
            if (string.IsNullOrWhiteSpace(width))
                return "0";

            if (type == TypeWidthTable.Pct)
                width = (double.Parse(width) / 50).ToString();
            else if (type == TypeWidthTable.Cm)
                width = (double.Parse(width) / Configuration.WidthCm).ToString();
            else
            {

            }

            return width;
        }
    }
}
