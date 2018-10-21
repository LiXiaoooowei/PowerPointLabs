using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.CaptionsLab
{
    public static class AnimationUtil
    {
        public static void ResetAnimationsForCalloutsOnSlide(Callouts callouts, PowerPointSlide s)
        {
            List<Shape> shapes = s.GetShapesWithPrefix("PowerPointLabs Callout ");
            s.RemoveAnimationsForShapes(shapes);
            Shape previous = null;
            for (int i = 0; i < callouts.GetNotesInvertedCount(); i++)
            {
                int shapeNo = callouts.GetCalloutIdxFromStmtNo(i);
                Shape captionBox = s.GetShapeWithName("PowerPointLabs Callout " + shapeNo)[0];
                if (i == 0)
                {
                    s.SetShapeAsAutoplay(captionBox);
                }

                if (i != 0)
                {
                    s.ShowShapeAfterClick(captionBox, i);
                    s.HideShapeAfterClick(previous, i);
                }

                if (i == callouts.GetNotesInvertedCount() - 1)
                {
                    s.HideShapeAsLastClickIfNeeded(captionBox);
                }
                previous = captionBox;
            }
        }

        public static Callouts SyncAnimationsForCalloutsOnSlide(PowerPointSlide slide, Callouts callouts)
        {
            Sequence sequence = slide.TimeLine.MainSequence;
            int clickNo = -1;
            Effect effect = sequence.FindFirstAnimationForClick(++clickNo);
            Dictionary<int, int> shapesAnime = new Dictionary<int, int>();
            List<int> shapesOrder = new List<int>();
            if (effect == null)
            {
                effect = sequence.FindFirstAnimationForClick(++clickNo);
            }
            while (effect != null)
            {
                int idx = GetShapeIdxForCallouts(effect.Shape.Name);
                Logger.Log("idx extracted is "+idx.ToString());
                if (idx != -1)
                {
                    if (!shapesOrder.Contains(idx))
                    {
                        Logger.Log("inserting shape with idx " + idx.ToString());
                        shapesAnime[idx] = clickNo;
                        shapesOrder.Add(idx);
                    }
                }
                effect = sequence.FindFirstAnimationForClick(++clickNo);
            }
            callouts.ReorderNotes(shapesOrder);
            ResetAnimationsForCalloutsOnSlide(callouts, slide);

            return callouts;
        }

        private static int GetShapeIdxForCallouts(string shapeName)
        {

            try
            {
                Match m = Regex.Match(shapeName, @"^PowerPointLabs\sCallout\s([1-9][0-9]*)$");
                if (m.Success && Int32.TryParse(m.Groups[1].Value, out int shapeNo))
                {
                    return shapeNo;
                }
                return -1;
            }
            catch (Exception)
            {
                return -1;
            }
        }
    }
}
