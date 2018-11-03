using System;
using System.Diagnostics;
using System.Globalization;

using PowerPointLabs.NarrationsLab.Data;
using PowerPointLabs.NarrationsLab.Views;

namespace PowerPointLabs.NarrationsLab.ValueConverters
{
    public class NarrationsLabSettingsPageValueConverter: BaseValueConverter<NarrationsLabSettingsPageValueConverter>
    {
        public override object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            Debug.WriteLine("Inside converter");
            // Find the appropriate page
            switch ((NarrationsLabSettingsPage)value)
            {
                case NarrationsLabSettingsPage.MainSettingsPage:
                    Debug.WriteLine("reached main settings page");
                    return NarrationsLabMainSettingsPage.GetInstance();
                case NarrationsLabSettingsPage.LoginPage:
                    return HumanVoiceLoginPage.GetInstance();
                case NarrationsLabSettingsPage.VoiceSelectionPage:
                    return HumanVoiceSelectionPage.GetInstance();
                default:
                    Debugger.Break();
                    return null;
            }
        }

        public override object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return null;
        }
    }
}
