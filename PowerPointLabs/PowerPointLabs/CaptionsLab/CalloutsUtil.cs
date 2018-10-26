using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.TagMatchers;
using PowerPointLabs.Tags;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;

namespace PowerPointLabs.CaptionsLab
{
    public static class CalloutsUtil
    {
#pragma warning disable 0618
        public static IEnumerable<string> SplitNotesByClicks(string rawNotes)
        {
            TaggedText taggedNotes = new TaggedText(rawNotes);
            List<string> splitByClicks = taggedNotes.SplitByClicks();
            return splitByClicks;
        }

        public static List<Tuple<NameTag, string>> ConvertNotesToCallouts(IEnumerable<string> separatedNotes)
        {
            List<Tuple<NameTag, string>> captionCollection = new List<Tuple<NameTag, string>>();
            HashSet<NameTag> tagSet = new HashSet<NameTag>();
            foreach (string text in separatedNotes)
            {
                string note = text;
                var match = Regex.Match(note, @"\[Name\s*:(.*)\]", RegexOptions.IgnoreCase);
                if (!match.Success)
                {
                    string uniqueTag = NameTagsUtil.GenerateUniqueName();
                    note = "[Name: " + uniqueTag + "]" + note; 
                }
                TaggedText section = new TaggedText(note);
                string currentCaption = section.ToPrettyString().Trim();
                List<NameTag> tags = new NameTagMatcher().NameTagMatches(note);
                if (!string.IsNullOrEmpty(currentCaption) && tags.Count == 1 && !tagSet.Contains(tags[0]))
                {
                    captionCollection.Add(new Tuple<NameTag, string>(tags[0], currentCaption));
                }
                else
                {
                    //TODO: Exception Handling
                }
            }
            return captionCollection;
        }

        public static List<string> ConvertNotesToCaptions(IEnumerable<string> separatedNotes)
        {
            List<string> captionCollection = new List<string>();
            foreach (string text in separatedNotes)
            {
                TaggedText section = new TaggedText(text);
                string currentCaption = section.ToPrettyString().Trim();
                if (!string.IsNullOrEmpty(currentCaption))
                {
                    captionCollection.Add(currentCaption);
                }
                else
                {
                    //TODO: Exception Handling
                }
            }
            return captionCollection;
        }
        public static string GetCalloutNotes(PowerPointSlide s)
        {
            StringBuilder builder = new StringBuilder();
            if (s.NotesPageText.Trim() != "")
            {
                builder.Append(s.NotesPageText.Trim());
            }
            try
            {
                ShapeRange shapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;

                foreach (Shape shape in shapes)
                {
                    if (!shape.Name.Contains("PPTLabs Callout "))
                    {
                        string newNote = shape.TextFrame.TextRange.Text;
                        string uniqueTag = NameTagsUtil.GenerateUniqueName();
                        shape.Name = "PPTLabs Callout " + uniqueTag;
                        builder.Append("[Name:" + uniqueTag + "]" + newNote + "[AfterClick]");
                    }
                }
            }
            catch (Exception e)
            {
                Logger.Log(e.Message);
            }
            return builder.ToString();
        }

        public static void UpdateCalloutBoxOnSlide(IntermediateResultTable intermediateResult, PowerPointSlide s)
        {
            foreach (Tuple<NameTag, string> note in intermediateResult.GetInsertedNotes())
            {
                Logger.Log("inserted note is " + note.Item2);
                InsertCalloutBoxToSlide(note.Item1, note.Item2, s);
            }

            foreach (Tuple<NameTag, string> note in intermediateResult.GetDeletedNotes())
            {
                Logger.Log("deleted note is " + note.Item2);
                DeleteCalloutBoxFromSlide(note.Item1, note.Item2, s);
            }

            foreach (Tuple<NameTag, string> note in intermediateResult.GetModifiedNotes())
            {
                Logger.Log("modified note is " + note.Item2);
                ModifyCalloutBoxFromSlide(note.Item1, note.Item2, s);
            }
        }

        private static void ModifyCalloutBoxFromSlide(NameTag tag, string note, PowerPointSlide s)
        {
            string shapeName = "PPTLabs Callout " + tag.Contents;
            List<Shape> shapes = s.GetShapeWithName(shapeName);
            if (shapes.Count != 0)
            {
                shapes[0].TextFrame.TextRange.Text = note;
            }
        }

        private static void DeleteCalloutBoxFromSlide(NameTag tag, string note, PowerPointSlide s)
        {
            string shapeName = "PPTLabs Callout " + tag.Contents;
            s.DeleteShapeWithName(shapeName);
        }

        private static Shape InsertCalloutBoxToSlide(NameTag tag, string note, PowerPointSlide s)
        {
            string shapeName = "PPTLabs Callout " + tag.Contents;
            if (s.HasShapeWithSameName(shapeName))
            {
                return null;
            }
            float slideWidth = PowerPointPresentation.Current.SlideWidth;
            float slideHeight = PowerPointPresentation.Current.SlideHeight;

            Shape callout = s.Shapes.AddCallout(MsoCalloutType.msoCalloutThree, 10, 10, 100, 10);
            callout.Name = shapeName;
            callout.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            callout.TextFrame.TextRange.Text = note;
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
