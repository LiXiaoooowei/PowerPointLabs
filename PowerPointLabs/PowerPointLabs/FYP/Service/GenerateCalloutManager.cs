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
    public class GenerateCalloutManager : LabCustomizationManager
    {
        public GenerateCalloutManager(string text, int tag, bool isActivated = false)
        {
            this.text = text;
            this.tag = tag;
            this.isActivated = isActivated;
        }
        public override List<Effect> PerformAction(PowerPointSlide slide, int clickNo, int seqNo = -1, bool isSeperateClick = false)
        {
            string name = FYPText.Identifier + FYPText.Underscore + tag.ToString() + FYPText.Underscore + FYPText.CalloutIdentifier;
            if (isActivated)
            {
                Shape shape = CalloutsUtil.InsertDefaultCalloutBoxToSlide(name, text, slide);
                Effect effectAppear = AnimationUtil.AppendAnimationsForCalloutsToSlide(shape, slide, clickNo);
                Effect effect = slide.SetShapeAsClickTriggered(shape, clickNo + 1, MsoAnimEffect.msoAnimEffectAppear, isSeperateClick);
                effect.Exit = Microsoft.Office.Core.MsoTriState.msoTrue;
                return new List<Effect>() { effectAppear };
            }
            else
            {
                slide.DeleteShapeWithName(name);
                return new List<Effect>() { };
            }
        }
    }
}
