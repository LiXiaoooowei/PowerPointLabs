﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Markup;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.FYP.Data;
using PowerPointLabs.Models;

namespace PowerPointLabs.FYP.Converters
{
#pragma warning disable 0618
    public class BlockItemIndexConverter : MarkupExtension, IValueConverter
    {

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            ListViewItem item = (ListViewItem)value;
            ListView listView = ItemsControl.ItemsControlFromItemContainer(item) as ListView;
            ObservableCollection<AnimationItem> items = listView.ItemsSource as ObservableCollection<AnimationItem>;
            int index = listView.ItemContainerGenerator.IndexFromContainer(item);
            AnimationItem animationItem = items.ElementAt(index);
            if (index == 0)
            {
                animationItem.ClickNo = PowerPointCurrentPresentationInfo.CurrentSlide.IsFirstAnimationTriggeredByClick() ? 1 : 0;
            }
            else if (animationItem is LabAnimationItem && (items.ElementAt(index - 1) is CustomAnimationItems) 
                && (animationItem as LabAnimationItem).IsTail)
            {
                animationItem.ClickNo = items.ElementAt(index - 1).ClickNo;
            }
            else
            {
               
                animationItem.ClickNo = items.ElementAt(index - 1).ClickNo + 1;
            }
            return animationItem.ClickNo;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return null;
        }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }
    }
}
