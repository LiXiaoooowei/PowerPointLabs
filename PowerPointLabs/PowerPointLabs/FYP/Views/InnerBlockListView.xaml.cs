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
        private ListViewDragDropManager<AnimationItem> innerDragManager;
        public InnerBlockListView()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            innerDragManager = new ListViewDragDropManager<AnimationItem>(listBox);
        }
    }
}
