using OpenXMLight.Configurations.Formatting;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Configurations.Elements.TableElements
{
    public class Borders : INotifyPropertyChanged
    {
        #region Private properties
        private BordersType lineType = BordersType.Single;
        private double lineWidth = 4;        
        #endregion

        public virtual BordersType LineType {
            get => lineType;
            set
            {
                lineType = value;

                OnPropertyChanged(nameof(LineType));
            } 
        }
        public virtual double LineWidth {
            get => lineWidth;
            set
            {
                lineWidth = value;

                OnPropertyChanged(nameof(LineWidth));
            }
        }


        public event PropertyChangedEventHandler? PropertyChanged;
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }
}
