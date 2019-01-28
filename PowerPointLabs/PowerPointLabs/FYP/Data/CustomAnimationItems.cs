using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.FYP.Data
{
    public class CustomAnimationItems:AnimationItem
    {
        public ObservableCollection<CustomAnimationItem> Items { get; set; }
        public CustomAnimationItems(ObservableCollection<CustomAnimationItem> items)
        {
            Items = items;
        }
    }
}
