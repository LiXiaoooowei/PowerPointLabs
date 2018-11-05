using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.CaptionsLab.CaptionsLabSettings.Storage;
using PowerPointLabs.CaptionsLab.CaptionsLabSettings.View;

namespace PowerPointLabs.CaptionsLab.CaptionsLabSettings
{
    internal static class CaptionsLabSettings
    {
        public static Shape shapeToCopy;
        public static void ShowSettingsDialog()
        {
            CaptionsLabSettingsDialogBox dialog = new CaptionsLabSettingsDialogBox();
            dialog.DialogConfirmedHandler += OnSettingsDialogConfirmed;
            dialog.ShowDialog();
        }
        private static void OnSettingsDialogConfirmed(string itemSource)
        {
            if (!string.IsNullOrWhiteSpace(ExtractShapeNameFromSource(itemSource)))
            {
                CaptionsLabPresentation pres = CaptionsLabPresentation.GetInstance(CaptionsLabStorageConfig.GetCalloutPptxStoragePath(), "new");
                shapeToCopy = pres.GetShapeWithName(ExtractShapeNameFromSource(itemSource));
            }
            if (shapeToCopy != null)
            {
                Logger.Log(shapeToCopy.Name);
            }
        }

        private static string ExtractShapeNameFromSource(string itemSource)
        {
            try
            {
                Match m = Regex.Match(itemSource, @"^.*\\([A-Za-z0-9\s]*).png$");
                if (m.Success)
                {
                    Logger.Log("name extracted is " + m.Groups[1].Value);
                    return m.Groups[1].Value;
                }
                return "";
            }
            catch (Exception)
            {
                return "";
            }
        }
    }
}
