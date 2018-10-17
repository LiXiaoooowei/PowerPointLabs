using MahApps.Metro.Controls;
using Microsoft.Office.Core;
using PowerPointLabs.Models;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.CaptionsLab
{
    /// <summary>
    /// Interaction logic for CalloutsTextDialog.xaml
    /// </summary>
    public partial class CalloutsTextDialog : MetroWindow
    {
#pragma warning disable 0618
        public CalloutsTextDialog()
        {
            InitializeComponent();
            DataContext = this;
        }
       
        public string Text { get; set; }
       
        private void Button_Click(object sender, System.Windows.RoutedEventArgs e)
        {
          //  AddCalloutToObject(Text);
            Close();
        }
    }
}
