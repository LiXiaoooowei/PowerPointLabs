using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.FYP
{
    [ExportLabelRibbonId("FYPLab")]
    class FYPLabLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return "FYP";
        }
    }
}
