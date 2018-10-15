using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.CaptionsLab.AddCallouts
{
    [ExportSupertipRibbonId(CaptionsLabText.AddCalloutsTag)]
    class AddCalloutsSupertipHandler: SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return CaptionsLabText.AddCalloutsButtonSupertip;
        }
    }
}
