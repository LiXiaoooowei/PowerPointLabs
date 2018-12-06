using Microsoft.Office.Tools;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.FYP;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.FYP
{
    [ExportActionRibbonId("FYPLab")]
    class FYPLabActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.RegisterTaskPane(typeof(FYPtaskpane), "FYP");
            CustomTaskPane fypTaskPane = this.GetTaskPane(typeof(FYPtaskpane));
            // if currently the pane is hidden, show the pane
            if (!fypTaskPane.Visible)
            {
                // fire the pane visble change event
                fypTaskPane.Visible = true;
            }
            else
            {
                fypTaskPane.Visible = false;
            }
        }
    }
}
