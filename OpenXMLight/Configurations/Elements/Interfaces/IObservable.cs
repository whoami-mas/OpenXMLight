using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Configurations.Elements.Interfaces
{
    public interface IObservable
    {
        //protected void RegisterObserver(IObserver o);
        //protected void RemoveObserver(IObserver o);
        void NotifyObservers();
    }
}
