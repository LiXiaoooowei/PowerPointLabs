using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.CaptionsLab.AddCallouts
{
    [ExportLabelRibbonId(CaptionsLabText.AddCalloutsTag)]
    class AddCalloutsLabelHandler: LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return CaptionsLabText.AddCalloutsButtonLabel;
        }
    }
}
