using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.FYP.Views
{
    /// <summary>
    /// Interaction logic for LabItemCreationCard.xaml
    /// </summary>
    public partial class LabItemCreationCard : UserControl
    {
        private ScrollBar _horizontalScrollBar;
        private RepeatButton _leftButton;
        private RepeatButton _rightButton;

        public LabItemCreationCard()
        {
            InitializeComponent();
        }

        private void HorizontalScrollViewer_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            ScrollViewer scrollViewer = sender as ScrollViewer;

            _horizontalScrollBar = scrollViewer.Template.FindName("PART_HorizontalScrollBar", scrollViewer) as ScrollBar;
            _leftButton = _horizontalScrollBar.Template.FindName("PART_LeftButton", _horizontalScrollBar) as RepeatButton;
            _rightButton = _horizontalScrollBar.Template.FindName("PART_RightButton", _horizontalScrollBar) as RepeatButton;

        }


        private void LeftButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            HorizontalScroller.LineLeft();
        }

        private void RightButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            HorizontalScroller.LineRight();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var siblings = ((sender as FrameworkElement).Parent as Panel).Children;
            var textBlock = siblings.OfType<TextBlock>().First();
            (itemsControl.ItemsSource as ObservableCollection<string>).Remove(textBlock.Text.ToString());
        }
    }
}
