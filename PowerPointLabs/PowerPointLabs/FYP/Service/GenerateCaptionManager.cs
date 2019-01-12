using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.CaptionsLab;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.FYP.Service
{
    public class GenerateCaptionManager : LabCustomizationManager
    {
        public GenerateCaptionManager(string text, int tag, bool isActivated = false)
        {
            this.text = text;
            this.tag = tag;
            this.isActivated = isActivated;
        }
        public override void PerformAction(PowerPointSlide slide, int clickNo, int seqNo = -1)
        {
            string name = FYPText.Identifier + FYPText.Underscore + tag.ToString() + FYPText.Underscore + FYPText.CaptionIdentifier;
            if (isActivated)
            {
                Shape shape = NotesToCaptions.AddCaptionBoxToSlide(text, name, slide);
                AnimationUtil.AppendAnimationsForCalloutsToSlide(shape, slide, clickNo);
                Effect effect = slide.SetShapeAsClickTriggered(shape, clickNo + 1, MsoAnimEffect.msoAnimEffectAppear);
                effect.Exit = Microsoft.Office.Core.MsoTriState.msoTrue;
            }
            else
            {
                slide.DeleteShapeWithName(name);
            }
        }
    }
}
