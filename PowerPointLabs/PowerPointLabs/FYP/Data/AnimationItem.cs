using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.FYP.Data
{
    public class AnimationItem: INotifyPropertyChanged
    {
        public int ClickNo
        {
            get
            {
                return clickNo;
            }
            set
            {
                if (clickNo != value)
                {
                    clickNo = value;
                    NotifyPropertyChanged();
                }
            }
        }
        private int clickNo;
        public AnimationItem(int clickNo)
        {
            this.clickNo = clickNo;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
