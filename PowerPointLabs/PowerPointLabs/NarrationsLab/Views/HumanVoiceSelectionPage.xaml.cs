using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
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

namespace PowerPointLabs.NarrationsLab.Views
{
    /// <summary>
    /// Interaction logic for HumanVoiceLoginPage.xaml
    /// </summary>
    public partial class HumanVoiceSelectionPage : Page
    {
        private static HumanVoiceSelectionPage instance;
        private HumanVoiceSelectionPage()
        {
            InitializeComponent();
        }

        public static HumanVoiceSelectionPage GetInstance()
        {
            if (instance == null)
            {
                instance = new HumanVoiceSelectionPage();
            }
            return instance;
        }
        public void Destroy()
        {
            instance = null;
        }
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            NarrationsLabSettingsDialogBox.GetInstance()
                .SetCurrentPage(DataModels.NarrationsLabSettingsPage.MainSettingsPage);
        }

        private void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine("setting human voice selected");
            NarrationsLabMainSettingsPage.GetInstance().SetHumanVoiceSelected("test");
            NarrationsLabSettingsDialogBox.GetInstance()
                .SetCurrentPage(DataModels.NarrationsLabSettingsPage.MainSettingsPage);
        }
    }
}
