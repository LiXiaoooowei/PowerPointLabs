using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

using Microsoft.Office.Core;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.CaptionsLab
{
    internal static class NotesToCallouts
    {
#pragma warning disable 0618
        public static void AddCallouts(float pLeft, float pTop, Shape triggerShape)
        {
            Globals.ThisAddIn.Ribbon.CalloutsTextDialog = new CalloutsTextDialog();
            Globals.ThisAddIn.Ribbon.CalloutsTextDialog.ShowDialog();
            string text = Globals.ThisAddIn.Ribbon.CalloutsTextDialog.Text;
            Shape callout = AddCalloutToObject(text, pLeft, pTop);
            ApplyAnimationToCallout(callout, triggerShape);
        }
        private static Shape AddCalloutToObject(string content, float pLeft, float pTop)
        {
            PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
            Shape callout = currentSlide.Shapes.AddCallout(MsoCalloutType.msoCalloutOne, pLeft, pTop, 100, 30);
            callout.TextFrame.TextRange.Text = content;
            callout.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            callout.TextFrame.WordWrap = MsoTriState.msoTrue;
            callout.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentCentered;
            callout.Top = pTop - callout.Height;
            return callout;
        }
        private static void ApplyAnimationToCallout(Shape s, Shape triggerShape)
        {
            PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
            currentSlide.TimeLine.InteractiveSequences.Add(1).AddTriggerEffect(s, PowerPoint.MsoAnimEffect.msoAnimEffectBoomerang, 
                PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnShapeClick, triggerShape);
        }
    }
}
