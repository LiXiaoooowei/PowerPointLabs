using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

using PowerPointLabs.NarrationsLab.Data;

namespace PowerPointLabs.NarrationsLab.Views
{
    /// <summary>
    /// Interaction logic for HumanVoiceLoginPage.xaml
    /// </summary>
    public partial class HumanVoiceSelectionPage : Page
    {
        private static HumanVoiceSelectionPage instance;
        private ObservableCollection<HumanVoice> voices = HumanVoiceList.voices;
        private HumanVoiceSelectionPage()
        {
            InitializeComponent();
            voiceList.ItemsSource = voices;
            voiceList.DisplayMemberPath = "Voice";
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
                .SetCurrentPage(Data.NarrationsLabSettingsPage.MainSettingsPage);
        }

        private void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            HumanVoice item = (HumanVoice)voiceList.SelectedItem;
            NarrationsLabSettings.humanVoice = item;
            NarrationsLabMainSettingsPage.GetInstance().SetHumanVoiceSelected(item.Voice.ToString());
            NarrationsLabSettingsDialogBox.GetInstance()
                .SetCurrentPage(Data.NarrationsLabSettingsPage.MainSettingsPage);
        }
    }
}
