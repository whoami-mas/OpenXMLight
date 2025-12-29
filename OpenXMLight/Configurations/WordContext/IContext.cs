using OpenXMLight.Configurations.Parts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Configurations.WordContext
{
    internal interface IContext
    {
        public Styles Styles { get; init; }
        public Endnotes Endnotes { get; init; }
    }
}
