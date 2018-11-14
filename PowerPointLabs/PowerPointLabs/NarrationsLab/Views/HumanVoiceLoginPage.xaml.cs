using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;

using PowerPointLabs.NarrationsLab.Data;
using PowerPointLabs.NarrationsLab.ViewModel;

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
            string _endpoint = endpoint.Text + "/issueToken";

            try
            {
                Authentication auth = new Authentication(_endpoint, _key);
                string accessToken = auth.GetAccessToken();
                Console.WriteLine("Token: {0}\n", accessToken);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed authentication.");
                Console.WriteLine(ex.ToString());
                Console.WriteLine(ex.Message);
                MessageBox.Show("Failed authentication");
                return;
            }
            
            UserAccount.GetInstance().SetUserKeyAndEndpoint(_key, _endpoint);
            NarrationsLabSettingsDialogBox.GetInstance()
                .SetCurrentPage(Data.NarrationsLabSettingsPage.VoiceSelectionPage);
        }
    }
}
