using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.CaptionsLab.CaptionsLabSettings.View;

namespace PowerPointLabs.CaptionsLab.CaptionsLabSettings
{
    internal static class CaptionsLabSettings
    {
        public static void ShowSettingsDialog()
        {
            CaptionsLabSettingsDialogBox dialog = new CaptionsLabSettingsDialogBox();
            dialog.ShowDialog();
        }
    }
}
