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
using PowerPointLabs.Tags;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.CaptionsLab
{
    internal static class AddCallouts
    {
#pragma warning disable 0618
        private static PowerPointCalloutsCache cache = PowerPointCalloutsCache.Instance;

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
                NameTagGenerator.GetTagNo(slide.NotesPageText.Trim());
                string contents = CalloutsUtil.GetCalloutNotes(slide);
                if (String.IsNullOrEmpty(contents) && !cache.IsTableExists(slide.Name))
                {
                    Logger.Log(String.Format("{0} in EmbedCaptionsOnSlides", CaptionsLabText.ErrorNoNotesLog));
                    MessageBox.Show(CaptionsLabText.ErrorNoNotes, CaptionsLabText.ErrorDialogTitle);
                    NotesToCaptions.ShowNotesPane();
                }              
                IEnumerable<string> splittedNotes = CalloutsUtil.SplitNotesByClicks(contents);
                List<Tuple<NameTag, string>> notes = CalloutsUtil.ConvertNotesToCallouts(splittedNotes);
                IntermediateResultTable intermediateResult = cache.UpdateNotes(slide.Name, notes);
                CalloutsUtil.UpdateCalloutBoxOnSlide(intermediateResult, slide);
                AnimationUtil.UpdateAnimationsForCalloutsOnSlide(intermediateResult, slide);
                slide.NotesPageText = intermediateResult.GetResultNotes();
            }
        }       
        
    }
}

