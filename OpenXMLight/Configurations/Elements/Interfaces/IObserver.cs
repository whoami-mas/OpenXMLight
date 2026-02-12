using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Configurations.Elements.Interfaces
{
    public interface IObserver
    {
        internal bool IsInitializedCache { get; set; }
        void RefreashCached();
    }
}
