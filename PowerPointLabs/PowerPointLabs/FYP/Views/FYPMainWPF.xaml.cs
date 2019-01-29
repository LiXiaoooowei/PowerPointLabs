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


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (tabControl.SelectedIndex == 0)
            {
                blockView.HandleSyncButtonClick();
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            string text = textBox.Text.Trim();
            if (text == "")
            {
                MessageBox.Show("Notes must not be empty!");
                return;
            }
            if (tabControl.SelectedIndex == 0)
            {
                blockView.AddLabAnimationItem(
                    new Data.LabAnimationItem(-1, text, LabAnimationItemIdentifierManager.GenerateUniqueNumber()));
            }

            textBox.Text = "";
        }
    }
}
