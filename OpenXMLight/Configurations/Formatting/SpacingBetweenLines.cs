using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Configurations.Formatting
{
    public class SpacingBetweenLines : INotifyPropertyChanged
    {
        //private int after = 100;
        //private int before = 100;
        //private int line = 200;
        private int after = 5;
        private int before = 5;
        private int line = 12;

        public int After
        {
            get => after * Configuration.InchInPixels;
            set
            {
                after = value;

                OnPropertyChanged(nameof(After));
            }
        }
        public int Before
        {
            get => before * Configuration.InchInPixels;
            set
            {
                before = value;

                OnPropertyChanged(nameof(Before));
            }
        }
        public int Line
        {
            get => line * Configuration.InchInPixels;
            set
            {
                line = value;

                OnPropertyChanged(nameof(Line));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            if(PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }
}
