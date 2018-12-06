using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.FYP
{
    [ExportSupertipRibbonId("FYPLab")]
    class FYPLabSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return PositionsLabText.RibbonMenuSupertip;
        }
    }
}
