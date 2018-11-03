using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportActionRibbonId(ShortcutsLabText.AddToCalloutShapeStyleTag)]
    class AddToCalloutShapeStyleActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            Selection selection = this.GetCurrentSelection();
            Shape selectedShape = selection.ShapeRange[1];
            if (selection.HasChildShapeRange)
            {
                selectedShape = selection.ChildShapeRange[1];
            }
            Logger.Log("clicked add to callout style");
        }
    }
}
