using OpenXMLight.Configurations.Formatting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Configurations.Elements.Interfaces
{
    public interface IMargin
    {
        public string Width { get; set; }
        public TypeWidthTable Type { get; }
    }
}
