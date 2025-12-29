using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.config
{
    internal interface IEnumValue<T>
    {
        T Value { get; }
    }
}
