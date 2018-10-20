using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using PowerPointLabs.Models;

namespace PowerPointLabs.CaptionsLab
{
    public static class CaptionUtil
    {
        public static IEnumerable<string> SplitNotesByClicks(string rawNotes)
        {
            TaggedText taggedNotes = new TaggedText(rawNotes);
            List<String> splitByClicks = taggedNotes.SplitByClicks();
            return splitByClicks;
        }

        public static List<string> ConvertSectionsToCaptions(IEnumerable<string> separatedNotes)
        {
            List<String> captionCollection = new List<string>();
            foreach (string text in separatedNotes)
            {
                TaggedText section = new TaggedText(text);
                String currentCaption = section.ToPrettyString().Trim();
                if (!string.IsNullOrEmpty(currentCaption))
                {
                    captionCollection.Add(currentCaption);
                }
            }
            return captionCollection;
        }

        public static bool IsNewNoteInserted(string s)
        {
            return s.Contains("[i]");
        }

        public static bool IsOldNoteDeleted(string s)
        {
            return s.Contains("[d]");
        }
    }
}
