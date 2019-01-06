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
        public override void PerformAction(PowerPointSlide slide, bool byClick = false)
        {
            if (isActivated)
            {
                Shape shape = CalloutsUtil.InsertDefaultCalloutBoxToSlide(
                    FYPText.Identifier + FYPText.Underscore + tag.ToString() + FYPText.Underscore + FYPText.CalloutIdentifier,
                    text, slide);
                AnimationUtil.AppendAnimationsForCalloutsToSlide(shape, slide, byClick);
            }
        }
    }
}
