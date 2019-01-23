using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.NarrationsLab.Data;
using PowerPointLabs.NarrationsLab.ViewModel;
using PowerPointLabs.SpeechEngine;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.NarrationsLab.Views
{
    public partial class VoicePreviewPage : Page
    {
        public VoicePreviewPage()
        {
            InitializeComponent();

            if (UserAccount.GetInstance().IsEmpty() || !IsValidUserAccount())
            {
                voiceList.Visibility = Visibility.Collapsed;
                humanVoiceBtn.Visibility = Visibility.Visible;
                changeAcctBtn.Visibility = Visibility.Hidden;
                RadioHumanVoice.IsEnabled = false;

            }
            else
            {
                string _key = UserAccount.GetInstance().GetKey();
                string _endpoint = UserAccount.GetInstance().GetRegion();

                voiceList.Visibility = Visibility.Visible;
                humanVoiceBtn.Visibility = Visibility.Collapsed;
                changeAcctBtn.Visibility = Visibility.Visible;
                RadioHumanVoice.IsEnabled = true;
            }
            voiceList.ItemsSource = voices;
            voiceList.DisplayMemberPath = "Voice";
        }
        public delegate void DialogConfirmedDelegate(string voiceName, HumanVoice humanVoiceName, bool isHumanVoiceSelected, bool preview);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }
        public string voicePreviewLabel;

        private static VoicePreviewPage instance;

        private ObservableCollection<HumanVoice> voices = HumanVoiceList.voices;

        public static VoicePreviewPage GetInstance()
        {
            if (instance == null)
            {
                instance = new VoicePreviewPage();
                instance.voicePreviewLabel = "";
            }
            else
            {
                if (UserAccount.GetInstance().IsEmpty())
                {
                    instance.voiceList.Visibility = Visibility.Collapsed;
                    instance.humanVoiceBtn.Visibility = Visibility.Visible;
                    instance.changeAcctBtn.Visibility = Visibility.Hidden;
                    instance.RadioHumanVoice.IsEnabled = false;
                }
                else
                {
                    instance.voiceList.Visibility = Visibility.Visible;
                    instance.humanVoiceBtn.Visibility = Visibility.Collapsed;
                    instance.changeAcctBtn.Visibility = Visibility.Visible;
                    instance.RadioHumanVoice.IsEnabled = true;
                }
            }
            return instance;
        }

        public void SetVoicePreviewSettings(int selectedVoiceIndex, HumanVoice humanVoice, List<string> voices, bool isHumanVoiceSelected, bool isPreviewChecked)
        {
            voiceSelectionInput.ItemsSource = voices;
            voiceSelectionInput.ToolTip = NarrationsLabText.SettingsVoiceSelectionInputTooltip;
            voiceSelectionInput.Content = voices[selectedVoiceIndex];

            if (humanVoice != null)
            {
                voiceList.SelectedItem = humanVoice;
            }

            RadioHumanVoice.IsChecked = isHumanVoiceSelected;
            RadioDefaultVoice.IsChecked = !isHumanVoiceSelected;

        }

        public NarrationVoices NarrationVoiceSelected
        {
            get
            {
                if (RadioDefaultVoice.IsChecked == true)
                {
                    return NarrationVoices.SystemVoice;
                }
                else
                {
                    return NarrationVoices.AzureVoice;
                }
            }
        }

        public void Destroy()
        {
            instance = null;
        }

        private bool IsValidUserAccount()
        {
            string _key = UserAccount.GetInstance().GetKey();
            string _endpoint = EndpointToUriMapping.regionToEndpointMapping[UserAccount.GetInstance().GetRegion()];

            try
            {
                Authentication auth = Authentication.GetInstance(_endpoint, _key);
                string accessToken = auth.GetAccessToken();
                Console.WriteLine("Token: {0}\n", accessToken);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed authentication.");
                Console.WriteLine(ex.ToString());
                Console.WriteLine(ex.Message);
                MessageBox.Show("Failed authentication");
                return false;
            }
            return true;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            
            string defaultVoiceSelected = RadioDefaultVoice.IsChecked == true ? voiceSelectionInput.Content.ToString() : null;
            HumanVoice humanVoiceSelected = RadioHumanVoice.IsChecked == true ? (HumanVoice)voiceList.SelectedItem : null;        
            string humanVoiceSelectedStr = humanVoiceSelected != null ? humanVoiceSelected.Voice.ToString() : null;
            instance.voicePreviewLabel = defaultVoiceSelected == null ? humanVoiceSelectedStr : defaultVoiceSelected;
            NarrationsLabSettingsDialogBox.GetInstance().DialogResult = true;
            NarrationsLabSettingsDialogBox.GetInstance().Close();
            NarrationsLabSettingsDialogBox.GetInstance().Destroy();
        }

        void VoiceSelectionInput_Item_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left && voiceSelectionInput.IsExpanded)
            {
                string value = ((TextBlock)e.Source).Text;
                voiceSelectionInput.Content = value;
            }
        }

        private void HumanVoiceBtn_Click(object sender, RoutedEventArgs e)
        {
            NarrationsLabSettingsDialogBox.GetInstance().SetCurrentPage(NarrationsLabSettingsPage.LoginPage);
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            NarrationsLabSettingsDialogBox.GetInstance().Destroy();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            NarrationsLabSettingsDialogBox.GetInstance().SetCurrentPage(NarrationsLabSettingsPage.LoginPage);
        }

        private void RadioDefaultVoice_Checked(object sender, RoutedEventArgs e)
        {
            RadioHumanVoice.IsChecked = false;
        }

        private void RadioHumanVoice_Checked(object sender, RoutedEventArgs e)
        {
            RadioDefaultVoice.IsChecked = false;
        }

        private void PreviewButton_Click(object sender, RoutedEventArgs e)
        {
            string textToSpeak = spokenText.Text.Trim();
            if (string.IsNullOrEmpty(textToSpeak))
            {
                return;
            }
            switch (NarrationVoiceSelected)
            {
                case NarrationVoices.SystemVoice:
                    string voiceName = voiceSelectionInput.Content.ToString();
                    TextToSpeech.SpeakString(textToSpeak, voiceName);
                    break;
                case NarrationVoices.AzureVoice:
                    HumanVoice humanVoiceSelected = RadioHumanVoice.IsChecked == true ? (HumanVoice)voiceList.SelectedItem : null;
                    TextToSpeech.SpeakTextWithAzureVoice(textToSpeak, humanVoiceSelected);
                    break;
                default:
                    break;
            }
        }
    }
}
