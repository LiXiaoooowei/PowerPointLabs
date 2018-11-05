using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
using PowerPointLabs.CaptionsLab.CaptionsLabSettings.Data;
using PowerPointLabs.CaptionsLab.CaptionsLabSettings.Storage;

namespace PowerPointLabs.CaptionsLab.CaptionsLabSettings.View
{
    /// <summary>
    /// Interaction logic for CaptionsLabSettingsDialogBox.xaml
    /// </summary>
    public partial class CaptionsLabSettingsDialogBox
    {
        public delegate void DialogConfirmedDelegate(string itemSource);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        private ShapeStyleList _calloutItems;
        public CaptionsLabSettingsDialogBox()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            ShapeStyle item = (ShapeStyle)listbox.SelectedItem;
            DialogConfirmedHandler(item.Source);
            Close();
        }

        private void MetroWindow_Loaded(object sender, RoutedEventArgs e)
        {
            _calloutItems = (ShapeStyleList)(Resources["CalloutItems"] as ObjectDataProvider).Data;
            _calloutItems.Path = CaptionsLabStorageConfig.GetCalloutImageStoragePath();
        }
    }
}
