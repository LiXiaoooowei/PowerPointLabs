using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.CaptionsLab;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.CaptionsLab
{
    [ExportActionRibbonId(CaptionsLabText.AddCalloutsTag)]
    class AddCalloutsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            //TODO: This needs to improved to stop using global variables
            this.StartNewUndoEntry();

            AddCallouts.EmbedCalloutsOnSelectedSlides(this.GetSelectedSlides());
            this.GetRibbonUi().RefreshRibbonControl("RemoveCaptionsButton");
        }
    }
}
