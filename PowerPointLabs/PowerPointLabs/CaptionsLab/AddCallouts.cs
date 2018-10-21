using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;

namespace PowerPointLabs.CaptionsLab
{
    internal static class AddCallouts
    {
#pragma warning disable 0618
        private static Dictionary<int, Callouts> calloutsTable = new Dictionary<int, Callouts>();

        public static void InitializeCalloutsTable(CalloutsTableSerializable table)
        {
            calloutsTable = table.ToCalloutsTable();
        }

        public static void EmbedCalloutsOnSelectedSlides(IEnumerable<PowerPointSlide> slides)
        {
            if (slides == null || !slides.Any())
            {
                Logger.Log(String.Format("{0} in EmbedCalloutsOnSelectedSlides", CaptionsLabText.ErrorNoSelectionLog));
                MessageBox.Show(CaptionsLabText.ErrorNoSelection, CaptionsLabText.ErrorDialogTitle);
                return;
            }
            EmbedCalloutsOnSlides(slides.ToList());
        }

        public static void EmbedCalloutsOnSlides(List<PowerPointSlide> slides)
        {
            foreach (PowerPointSlide slide in slides)
            {
                if (NeedsUpdateCallouts(slide.Name))
                {
                    InsertCalloutsOnSlideToNotesPage(slide);
                    UpdateCalloutsOnNotesPageToSlide(slide);
                }
                else
                {
                    bool captionAdded = EmbedCalloutsOnNotesPageToSlide(slide);
                    bool captionInserted = InsertCalloutsOnSlideToNotesPage(slide);
                    if (!captionAdded && slides.Count == 1 && !captionInserted)
                    {
                        Logger.Log(String.Format("{0} in EmbedCaptionsOnSlides", CaptionsLabText.ErrorNoNotesLog));
                        MessageBox.Show(CaptionsLabText.ErrorNoNotes, CaptionsLabText.ErrorDialogTitle);
                        NotesToCaptions.ShowNotesPane();
                    }
                }
                AnimationUtil.ResetAnimationsForCalloutsOnSlide(calloutsTable[slide.GetSlideIndexForCallouts()], slide);
            }
        }

        public static void SyncCalloutsOnSlides(List<PowerPointSlide> slides)
        {
            foreach (PowerPointSlide slide in slides)
            {
                int slideNo = slide.GetSlideIndexForCallouts();
                if (slideNo != -1 && calloutsTable.ContainsKey(slideNo))
                {
                    SyncCalloutsOnSlideToNotespage(slide);
                    calloutsTable[slideNo] = AnimationUtil.SyncAnimationsForCalloutsOnSlide(slide, calloutsTable[slideNo]);
                    slide.NotesPageText = calloutsTable[slideNo].NotesToString();
                }
                Logger.Log("saving slide " + slide.Name);
            }
            string filePath = Environment.ExpandEnvironmentVariables(@"%UserProfile%\\Desktop\\callouts.dat");
            StorageUtil.WriteToXMLFile(filePath, ConvertToCalloutTableSerializable());
        }

        private static CalloutsTableSerializable ConvertToCalloutTableSerializable()
        {
            List<CalloutsTableSerializable.CalloutsListSerializable> lists =
                new List<CalloutsTableSerializable.CalloutsListSerializable>();
            foreach (KeyValuePair<int, Callouts> callout in calloutsTable)
            {
                CalloutsTableSerializable.CalloutsListSerializable item =
                    new CalloutsTableSerializable.CalloutsListSerializable()
                    {
                        slideNo = callout.Key,
                        callout = callout.Value.Serialize()
                    };
                lists.Add(item);
            }
            return new CalloutsTableSerializable()
            {
                list = lists
            };
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

        private static bool InsertCalloutsOnSlideToNotesPage(PowerPointSlide s)
        {
            int slideNo = s.GetSlideIndexForCallouts();
            if (slideNo == -1)
            {
                slideNo = calloutsTable.Count + 1;
                s.Name = "PowerPointSlide " + slideNo.ToString();
            }
            if (!calloutsTable.ContainsKey(slideNo))
            {
                calloutsTable.Add(slideNo, new Callouts());
            }
            try
            {
                ShapeRange shapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                foreach (Shape shape in shapes)
                {
                    if (!shape.Name.Contains("PowerPointLabs Callout"))
                    {
                        string currentCaption = shape.TextFrame.TextRange.Text;
                        int calloutIdx = calloutsTable[slideNo].InsertCallout(currentCaption);
                        shape.Name = "PowerPointLabs Callout " + calloutIdx;
                    }
                }
                s.NotesPageText = calloutsTable[slideNo].NotesToString();
                return true;
            }
            catch (Exception e)
            {
                Logger.Log("Exception here " + e.Message);
                return false;
            }
        }

        private static bool SyncCalloutsOnSlideToNotespage(PowerPointSlide s)
        {
            int slideNo = s.GetSlideIndexForCallouts();
            if (slideNo == -1)
            {
                return false;
            }

            int calloutsDeleted = 0;
            Callouts calloutsOnNotesPage = calloutsTable[slideNo];

            for (int i = 0; i < calloutsOnNotesPage.GetNotesInvertedCount(); i++)
            {
                int currIdx = i - calloutsDeleted;
                int calloutIdx = calloutsOnNotesPage.GetCalloutIdxFromStmtNo(currIdx);
                bool foundShape = false;
                foreach (Shape shape in s.GetShapesWithPrefix("PowerPointLabs Callout"))
                {
                    if (shape.Name == "PowerPointLabs Callout " + calloutIdx.ToString())
                    {
                        calloutsOnNotesPage.UpdateCallout(shape.TextFrame.TextRange.Text, currIdx);
                        foundShape = true;
                        break;
                    }
                }
                if (!foundShape)
                {
                    calloutsOnNotesPage.DeleteCallout(currIdx);
                }
            }
            s.NotesPageText = calloutsTable[slideNo].NotesToString();
            return true;
        }

        private static bool UpdateCalloutsOnNotesPageToSlide(PowerPointSlide s)
        {
            String rawNotes = s.NotesPageText;
            int slideNo = s.GetSlideIndexForCallouts();
            if (String.IsNullOrWhiteSpace(rawNotes) || slideNo == -1)
            {
                return false;
            }

            IEnumerable<string> separatedNotes = CaptionUtil.SplitNotesByClicks(rawNotes);
            List<string> captionCollection = CaptionUtil.ConvertSectionsToCaptions(separatedNotes);
            if (captionCollection.Count == 0)
            {
                return false;
            }
            int calloutsDeleted = 0;
            for (int i = 0; i < captionCollection.Count; i++)
            {
                int currIdx = i - calloutsDeleted;
                string currentCaption = captionCollection[i];
                if (CaptionUtil.IsNewNoteInserted(currentCaption))
                {
                    currentCaption = currentCaption.Replace("[i]", "");
                    Shape callout = AddCalloutBoxToSlide(currentCaption, s);
                    int calloutIdx = calloutsTable[slideNo].InsertCallout(currentCaption, currIdx);
                    callout.Name = "PowerPointLabs Callout " + calloutIdx;
                }
                else if (CaptionUtil.IsOldNoteDeleted(currentCaption))
                {
                    calloutsDeleted++;
                    int calloutNo = calloutsTable[slideNo].DeleteCallout(currIdx);
                    s.RemoveShapeWithName("PowerPointLabs Callout " + calloutNo.ToString());
                }
                else
                {
                    int calloutNo = calloutsTable[slideNo].UpdateCallout(currentCaption, currIdx);
                    foreach (Shape callout in s.GetShapeWithName("PowerPointLabs Callout " + calloutNo.ToString()))
                    {
                        callout.TextFrame.TextRange.Text = currentCaption;
                    }
                }
            }
            s.NotesPageText = calloutsTable[slideNo].NotesToString();
            return true;
        }

        private static bool EmbedCalloutsOnNotesPageToSlide(PowerPointSlide s)
        {
            String rawNotes = s.NotesPageText;
            int slideNo = calloutsTable.Count + 1;
            s.Name = "PowerPointSlide " + slideNo.ToString();
            if (String.IsNullOrWhiteSpace(rawNotes))
            {
                return false;
            }

            IEnumerable<string> separatedNotes = CaptionUtil.SplitNotesByClicks(rawNotes);
            List<string> captionCollection = CaptionUtil.ConvertSectionsToCaptions(separatedNotes);
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
            s.NotesPageText = calloutsTable[slideNo].NotesToString();
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

