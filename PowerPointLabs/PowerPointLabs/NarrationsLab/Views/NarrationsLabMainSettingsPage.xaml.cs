using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;


using PowerPointLabs.NarrationsLab.DataModels;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.NarrationsLab.Views
{
    /// <summary>
    /// Interaction logic for NarrationsLabMainSettingsPage.xaml
    /// </summary>
    public partial class NarrationsLabMainSettingsPage: Page
    {
        public delegate void DialogConfirmedDelegate(string voiceName, bool preview);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }
       
        private static NarrationsLabMainSettingsPage instance;

        private NarrationsLabMainSettingsPage()
        {
            InitializeComponent();
        }
        public static NarrationsLabMainSettingsPage GetInstance()
        {
            if (instance == null)
            {
                instance = new NarrationsLabMainSettingsPage();
            }
            return instance;
        }

        public void SetNarrationsLabMainSettings(int selectedVoiceIndex, List<string> voices, bool isPreviewChecked)
        {
            voiceSelectionInput.ItemsSource = voices;
            voiceSelectionInput.ToolTip = NarrationsLabText.SettingsVoiceSelectionInputTooltip;
            voiceSelectionInput.Content = voices[selectedVoiceIndex];

            previewCheckbox.IsChecked = isPreviewChecked;
            previewCheckbox.ToolTip = NarrationsLabText.SettingsPreviewCheckboxTooltip;
        }

        public void SetHumanVoiceSelected(string voice)
        {
            humanVoiceChosen.Text = voice;
        }

        public void Destroy()
        {
            instance = null;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogConfirmedHandler(voiceSelectionInput.Content.ToString(), previewCheckbox.IsChecked.GetValueOrDefault());
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
    }
}
