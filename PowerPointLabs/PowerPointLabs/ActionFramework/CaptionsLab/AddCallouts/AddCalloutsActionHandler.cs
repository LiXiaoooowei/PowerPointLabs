using System;
using System.Collections.Generic;
using System.Windows;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;

using PowerPointLabs.CaptionsLab;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ActionFramework.CaptionsLab.AddCallouts
{
    [ExportActionRibbonId(CaptionsLabText.AddCalloutsTag)]
    class AddCalloutsActionHandler: ActionHandler
    {
#pragma warning disable 0618
        protected override void ExecuteAction(string ribbonId)
        {
            PowerPoint.Selection selection = this.GetCurrentSelection();
            PowerPoint.ShapeRange selectedShapes;
            try
            {
                selectedShapes = selection.ShapeRange;
            }
            catch (Exception)
            {
                MessageBox.Show("Please select at least one object to add callouts!");
                return;
            }
            NotesToCallouts.AddCallouts(selectedShapes[1].Left, selectedShapes[1].Top, selectedShapes[1]);
        }
    }
}
