using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXMLight.Spreadsheet.Parts;

namespace OpenXMLight.Spreadsheet.ExcelContext
{
    internal interface IContext
    {
        public Styles Styles { get; init; }
        public SharedStrings SharedStrings { get; init; }
    }
}
