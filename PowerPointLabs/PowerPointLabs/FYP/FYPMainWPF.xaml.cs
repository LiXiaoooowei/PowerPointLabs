using System;
using System.Collections.Generic;
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
using PowerPointLabs.FYP.Service;
using PowerPointLabs.FYP.Views;

namespace PowerPointLabs.FYP
{
    /// <summary>
    /// Interaction logic for FYPMainWPF.xaml
    /// </summary>
    public partial class FYPMainWPF : UserControl
    {
        public FYPMainWPF()
        {
            InitializeComponent();
        }

        private void PPTLabsPageView_Loaded(object sender, RoutedEventArgs e)
        {

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (tabControl.SelectedIndex == 0)
            {
                blockView.HandleButtonClick();
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (tabControl.SelectedIndex == 0)
            {
                string note = textBox.Text.Trim();
                if (note == "")
                {
                    MessageBox.Show("Notes must not be empty!");
                    return;
                }               
                blockView.AddLabAnimationItem(
                    new Data.LabAnimationItem(note, LabAnimationItemIdentifierManager.GenerateUniqueNumber()));
            }
        }
    }
}
