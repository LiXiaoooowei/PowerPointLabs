using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.FYP.Data;
using PowerPointLabs.FYP.Service;

namespace PowerPointLabs.FYP.Views
{
    /// <summary>
    /// Interaction logic for InnerBlockListView.xaml
    /// </summary>
    public partial class InnerBlockListView : UserControl
    {
        public ObservableCollection<AnimationItem> Items { get; private set; }

        public static readonly RoutedEvent UpBtnClickedEvent = EventManager.RegisterRoutedEvent(
    "UpBtnClickedHandlerParent",
    RoutingStrategy.Bubble,
    typeof(RoutedEventHandler),
    typeof(InnerBlockListView));

        public static readonly RoutedEvent DownBtnClickedEvent = EventManager.RegisterRoutedEvent(
            "DownBtnClickedHandlerParent",
            RoutingStrategy.Bubble,
            typeof(RoutedEventHandler),
            typeof(InnerBlockListView));

        public event RoutedEventHandler UpBtnClickedHandlerParent
        {
            add { AddHandler(UpBtnClickedEvent, value); }
            remove { RemoveHandler(UpBtnClickedEvent, value); }
        }

        public event RoutedEventHandler DownBtnClickedHandlerParent
        {
            add { AddHandler(DownBtnClickedEvent, value); }
            remove { RemoveHandler(DownBtnClickedEvent, value); }
        }

        private ListViewDragDropManager<AnimationItem> innerDragManager;
        public InnerBlockListView()
        {
            InitializeComponent();
        }

        private void HandleUpBtnClickedEvent(object sender, RoutedEventArgs e)
        {
            LabAnimationItem item = ((Button)e.OriginalSource).CommandParameter as LabAnimationItem;
            ObservableCollection<AnimationItem> items = listBox.ItemsSource as ObservableCollection<AnimationItem>;
            int idx = items.IndexOf(item);
            if (idx == 0)
            {
                RoutedEventArgs eventArgs = new RoutedEventArgs(UpBtnClickedEvent);
                eventArgs.Source = e.OriginalSource;
                RaiseEvent(eventArgs);
            }
            else
            {
                items.Move(idx, idx - 1);
            }
        }

        private void HandleDownBtnClickedEvent(object sender, RoutedEventArgs e)
        {
            LabAnimationItem item = ((Button)e.OriginalSource).CommandParameter as LabAnimationItem;
            ObservableCollection<AnimationItem> items = listBox.ItemsSource as ObservableCollection<AnimationItem>;
            int idx = items.IndexOf(item);
            if (idx == items.Count - 1)
            {
                RoutedEventArgs eventArgs = new RoutedEventArgs(DownBtnClickedEvent);
                eventArgs.Source = e.OriginalSource;
                RaiseEvent(eventArgs);
            }
            else
            {
                items.Move(idx, idx + 1);
            }
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            innerDragManager = new ListViewDragDropManager<AnimationItem>(listBox);
        }

    }
}
