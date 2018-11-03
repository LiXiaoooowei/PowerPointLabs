using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;

using PowerPointLabs.NarrationsLab.Data;

namespace PowerPointLabs.NarrationsLab.Views
{
    /// <summary>
    /// Interaction logic for HumanVoiceLoginPage.xaml
    /// </summary>
    public partial class HumanVoiceLoginPage : Page
    {
        private static HumanVoiceLoginPage instance;
        private HumanVoiceLoginPage()
        {
            InitializeComponent();
        }

        public static HumanVoiceLoginPage GetInstance()
        {
            if (instance == null)
            {
                instance = new HumanVoiceLoginPage();
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
                .SetCurrentPage(Data.NarrationsLabSettingsPage.MainSettingsPage); 
        }

        private void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            string _key = key.Text;
            string _endpoint = endpoint.Text;

            UserAccount.GetInstance().SetUserKeyAndEndpoint(_key, _endpoint);

            NarrationsLabSettingsDialogBox.GetInstance()
                .SetCurrentPage(Data.NarrationsLabSettingsPage.VoiceSelectionPage);
        }
    }
}
