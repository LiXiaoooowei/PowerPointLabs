using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.CaptionsLab
{
    internal static class AddCallouts
    {
#pragma warning disable 0618
        private static Dictionary<int, Callouts> calloutsTable = new Dictionary<int, Callouts>();

        public static void EmbedCalloutsOnSelectedSlides()
        {
            if (PowerPointCurrentPresentationInfo.SelectedSlides == null ||
                !PowerPointCurrentPresentationInfo.SelectedSlides.Any())
            {
                Logger.Log(String.Format("{0} in EmbedCalloutsOnSelectedSlides", CaptionsLabText.ErrorNoSelectionLog));
                MessageBox.Show(CaptionsLabText.ErrorNoSelection, CaptionsLabText.ErrorDialogTitle);
                return;
            }
            EmbedCalloutsOnSlides(PowerPointCurrentPresentationInfo.SelectedSlides.ToList());
        }

        public static void EmbedCalloutsOnSlides(List<PowerPointSlide> slides)
        {
            foreach (PowerPointSlide slide in slides)
            {
                if (NeedsUpdateCallouts(slide.Name))
                {
                     Logger.Log("calling updatecalloutsonslide");
                     UpdateCalloutsOnSlide(slide);
                }
                else
                {
                    Logger.Log("calling embedcalloutsonslide");
                    bool captionAdded = EmbedCalloutsOnSlide(slide);
                    if (!captionAdded && slides.Count == 1)
                    {
                        Logger.Log(String.Format("{0} in EmbedCaptionsOnSlides", CaptionsLabText.ErrorNoNotesLog));
                        MessageBox.Show(CaptionsLabText.ErrorNoNotes, CaptionsLabText.ErrorDialogTitle);
                        NotesToCaptions.ShowNotesPane();
                    }
                }
            }
        }

        public static void UpdateCalloutsOnSlides(List<PowerPointSlide> slides)
        {
            foreach (PowerPointSlide slide in slides)
            {
              //  UpdateCalloutsOnSlide(slide);
                Logger.Log("saving slide " + slide.Name);
            }
        }

        private static bool NeedsUpdateCallouts(string name)
        {
            Match m;
            try
            {
                Logger.Log("slide name to check is " + name);
                m = Regex.Match(name, @"^PowerPointSlide\s([1-9][0-9]*)$");
                if (m.Success && Int32.TryParse(m.Groups[1].Value, out int slideNo))
                {
                    Logger.Log("m.Value = " + m.Groups[1].Value);
                    return calloutsTable.ContainsKey(slideNo);
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool UpdateCalloutsOnSlide(PowerPointSlide s)
        {
            String rawNotes = s.NotesPageText;
            int slideNo = s.GetSlideIndexForCallouts();
            if (String.IsNullOrWhiteSpace(rawNotes)||slideNo == -1)
            {
                return false;
            }

            IEnumerable<string> separatedNotes = NotesToCaptions.SplitNotesByClicks(rawNotes);
            List<string> captionCollection = NotesToCaptions.ConvertSectionsToCaptions(separatedNotes);
            if (captionCollection.Count == 0)
            {
                return false;
            }

            for (int i = 0; i < captionCollection.Count; i++)
            {
                string currentCaption = captionCollection[i];
                //TODO: consider stmt insertion and deletion case
                int calloutNo = calloutsTable[slideNo].GetCalloutNoFromStmtNo(i);
                foreach (Shape callout in s.GetShapeWithName("PowerPointLabs Callout " + calloutNo.ToString()))
                {
                    callout.TextFrame.TextRange.Text = currentCaption;
                }
            }

            return true;
        }

        // Returns true if the captions are successfully added
        private static bool EmbedCalloutsOnSlide(PowerPointSlide s)
        {
            String rawNotes = s.NotesPageText;
            int slideNo = calloutsTable.Count + 1;
            s.Name = "PowerPointSlide " + slideNo.ToString();
            if (String.IsNullOrWhiteSpace(rawNotes))
            {
                return false;
            }

            IEnumerable<string> separatedNotes = NotesToCaptions.SplitNotesByClicks(rawNotes);
            List<string> captionCollection = NotesToCaptions.ConvertSectionsToCaptions(separatedNotes);
            if (captionCollection.Count == 0)
            {
                return false;
            }

            Shape previous = null;
            Callouts callouts = new Callouts();
            for (int i = 0; i < captionCollection.Count; i++)
            {
                String currentCaption = captionCollection[i];
                Shape callout = AddCalloutBoxToSlide(currentCaption, s);
                int calloutIdx = callouts.InsertCallout(currentCaption, i);
                callout.Name = "PowerPointLabs Callout " + calloutIdx;
                
                if (i == 0)
                {
                    s.SetShapeAsAutoplay(callout);
                }

                if (i != 0)
                {
                    s.ShowShapeAfterClick(callout, i);
                    s.HideShapeAfterClick(previous, i);
                }

                if (i == captionCollection.Count - 1)
                {
                    s.HideShapeAsLastClickIfNeeded(callout);
                }
                previous = callout;
            }
            calloutsTable.Add(slideNo, callouts);
            return true;
        }

        private static Shape AddCalloutBoxToSlide(string caption, PowerPointSlide s)
        {
            float slideWidth = PowerPointPresentation.Current.SlideWidth;
            float slideHeight = PowerPointPresentation.Current.SlideHeight;

            Shape callout = s.Shapes.AddCallout(MsoCalloutType.msoCalloutThree, 10, 10, 100, 10);
            callout.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            callout.TextFrame.TextRange.Text = caption;
            callout.TextFrame.WordWrap = MsoTriState.msoTrue;
            callout.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentCentered;
            callout.TextFrame.TextRange.Font.Size = 12;
            callout.Fill.BackColor.RGB = 0;
            callout.Fill.Transparency = 0.2f;
            callout.TextFrame.TextRange.Font.Color.RGB = 0xffffff;

            return callout;
        }

    }
}

