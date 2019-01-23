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

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.FYP.Data;
using PowerPointLabs.NarrationsLab;
using PowerPointLabs.NarrationsLab.Storage;
using PowerPointLabs.NarrationsLab.Views;

namespace PowerPointLabs.FYP.Views
{
    /// <summary>
    /// Interaction logic for LabAnimationItemCard.xaml
    /// </summary>
    public partial class LabAnimationItemCard : UserControl
    {

        public static readonly RoutedEvent UpBtnClickedEvent = EventManager.RegisterRoutedEvent(
            "UpBtnClickedHandler", 
            RoutingStrategy.Bubble, 
            typeof(RoutedEventHandler), 
            typeof(LabAnimationItemCard));

        public static readonly RoutedEvent DownBtnClickedEvent = EventManager.RegisterRoutedEvent(
            "DownBtnClickedHandler",
            RoutingStrategy.Bubble,
            typeof(RoutedEventHandler),
            typeof(LabAnimationItemCard));

        public event RoutedEventHandler UpBtnClickedHandler
        {
            add { AddHandler(UpBtnClickedEvent, value); }
            remove { RemoveHandler(UpBtnClickedEvent, value); }
        }

        public event RoutedEventHandler DownBtnClickedHandler
        {
            add { AddHandler(DownBtnClickedEvent, value); }
            remove { RemoveHandler(DownBtnClickedEvent, value); }
        }

        public LabAnimationItemCard()
        {
            InitializeComponent();
        }

        private void UpBtn_Click(object sender, RoutedEventArgs e)
        {
            RoutedEventArgs eventArgs = new RoutedEventArgs(UpBtnClickedEvent);
            eventArgs.Source = sender;
            RaiseEvent(eventArgs);
        }

        private void DownBtn_Click(object sender, RoutedEventArgs e)
        {
            RoutedEventArgs eventArgs = new RoutedEventArgs(DownBtnClickedEvent);
            eventArgs.Source = sender;
            RaiseEvent(eventArgs);
        }

        private void VoicePreviewButton_Click(object sender, RoutedEventArgs e)
        {
            NarrationsLabStorageConfig.LoadUserAccount();

            NarrationsLabSettingsDialogBox dialog = NarrationsLabSettingsDialogBox.GetInstance(
                NarrationsLab.Data.NarrationsLabSettingsPage.VoicePreviewPage);
            dialog.Height = 250;
            VoicePreviewPage.GetInstance().SetVoicePreviewSettings(
               NarrationsLabSettings.VoiceSelectedIndex,
               NarrationsLabSettings.humanVoice,
               NarrationsLabSettings.VoiceNameList,
               NotesToAudio.IsHumanVoiceSelected,
               NarrationsLabSettings.IsPreviewEnabled);
            VoicePreviewPage.GetInstance().DialogConfirmedHandler += NarrationsLabSettings.OnSettingsDialogConfirmed;
            VoicePreviewPage.GetInstance().spokenText.Text = explanatoryNote.Text;

            if (dialog.ShowDialog() == true)
            {
                previewVoiceLabel.Content = VoicePreviewPage.GetInstance().voicePreviewLabel;
                explanatoryNote.Text = VoicePreviewPage.GetInstance().spokenText.Text;
            }
        }
    }
}
