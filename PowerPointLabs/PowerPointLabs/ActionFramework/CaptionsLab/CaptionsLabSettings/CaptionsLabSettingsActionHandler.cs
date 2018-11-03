using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.CaptionsLab.CaptionsLabSettings;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.NarrationsLab
{
    [ExportActionRibbonId(CaptionsLabText.SettingsTag)]
    class CaptionsLabSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            CaptionsLabSettings.ShowSettingsDialog();
        }
    }
}
