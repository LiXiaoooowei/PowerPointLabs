using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;


using PowerPointLabs.NarrationsLab.Data;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.NarrationsLab.Views
{
    /// <summary>
    /// Interaction logic for NarrationsLabMainSettingsPage.xaml
    /// </summary>
    public partial class NarrationsLabMainSettingsPage: Page
    {
        public delegate void DialogConfirmedDelegate(string voiceName, HumanVoice humanVoiceName, bool isHumanVoiceSelected, bool preview);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }
       
        private static NarrationsLabMainSettingsPage instance;

        private ObservableCollection<HumanVoice> voices = HumanVoiceList.voices;

        private NarrationsLabMainSettingsPage()
        {
            InitializeComponent();
            if (UserAccount.GetInstance().IsEmpty())
            {
                voiceList.Visibility = Visibility.Collapsed;
                humanVoiceBtn.Visibility = Visibility.Visible;
                changeAcctBtn.Visibility = Visibility.Hidden;
            }
            else
            {
                voiceList.Visibility = Visibility.Visible;
                humanVoiceBtn.Visibility = Visibility.Collapsed;
                changeAcctBtn.Visibility = Visibility.Visible;
            }
            voiceList.ItemsSource = voices;
            voiceList.DisplayMemberPath = "Voice";
        }
        public static NarrationsLabMainSettingsPage GetInstance()
        {
            if (instance == null)
            {
                instance = new NarrationsLabMainSettingsPage();
            }
            else
            {
                if (UserAccount.GetInstance().IsEmpty())
                {
                    instance.voiceList.Visibility = Visibility.Collapsed;
                    instance.humanVoiceBtn.Visibility = Visibility.Visible;
                    instance.changeAcctBtn.Visibility = Visibility.Hidden;
                }
                else
                {
                    instance.voiceList.Visibility = Visibility.Visible;
                    instance.humanVoiceBtn.Visibility = Visibility.Collapsed;
                    instance.changeAcctBtn.Visibility = Visibility.Visible;
                }
            }
            return instance;
        }

        public void SetNarrationsLabMainSettings(int selectedVoiceIndex, HumanVoice humanVoice, List<string> voices, bool isHumanVoiceSelected, bool isPreviewChecked)
        {
            voiceSelectionInput.ItemsSource = voices;
            voiceSelectionInput.ToolTip = NarrationsLabText.SettingsVoiceSelectionInputTooltip;
            voiceSelectionInput.Content = voices[selectedVoiceIndex];

            if (humanVoice != null)
            {
                voiceList.SelectedItem = humanVoice;
            }

            previewCheckbox.IsChecked = isPreviewChecked;
            previewCheckbox.ToolTip = NarrationsLabText.SettingsPreviewCheckboxTooltip;
          
                RadioHumanVoice.IsChecked = isHumanVoiceSelected;
                RadioDefaultVoice.IsChecked = !isHumanVoiceSelected;
            
        }

        public void Destroy()
        {
            instance = null;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            string defaultVoiceSelected = RadioDefaultVoice.IsChecked == true? voiceSelectionInput.Content.ToString() : null;
            HumanVoice humanVoiceSelected = RadioHumanVoice.IsChecked == true ? (HumanVoice)voiceList.SelectedItem : null;
            DialogConfirmedHandler(defaultVoiceSelected, humanVoiceSelected, humanVoiceSelected != null, previewCheckbox.IsChecked.GetValueOrDefault());
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
    }
}
