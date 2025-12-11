using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXMLight.Configurations.Elements.Table;

namespace OpenXMLight.Configurations
{
    public interface ICellObserver
    {
        void OnCellsMerged(HashSet<int> cells);
    }
}
