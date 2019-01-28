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

using PowerPointLabs.FYP.Data;

namespace PowerPointLabs.FYP.Views
{
    /// <summary>
    /// Interaction logic for CustomAnimationItemCard.xaml
    /// </summary>
    /// 
    public partial class CustomAnimationItemCard : UserControl
    {
        public ObservableCollection<CustomAnimationItem> CustomItems { get; private set; }
        public CustomAnimationItemCard()
        {
            InitializeComponent();
        }
    }
}
