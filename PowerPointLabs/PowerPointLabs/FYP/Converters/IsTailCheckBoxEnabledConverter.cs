using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Markup;

using PowerPointLabs.FYP.Data;

namespace PowerPointLabs.FYP.Converters
{
    public class IsTailCheckBoxEnabledConverter : MarkupExtension, IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            ListViewItem item = (ListViewItem)value;
            ListView listView = ItemsControl.ItemsControlFromItemContainer(item) as ListView;
            ObservableCollection<AnimationItem> items = listView.ItemsSource as ObservableCollection<AnimationItem>;
            int index = listView.ItemContainerGenerator.IndexFromContainer(item);
            bool isTailEnabled = index != 0 && (items.ElementAt(index - 1) is CustomAnimationItems);
            LabAnimationItem labItem = items.ElementAt(index) as LabAnimationItem;
            if (labItem != null)
            {
                labItem.IsTailEnabled = isTailEnabled;
            }
            return isTailEnabled;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }
    }
}
