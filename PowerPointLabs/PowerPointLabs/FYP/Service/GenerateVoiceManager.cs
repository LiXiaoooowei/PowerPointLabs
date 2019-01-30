using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using PowerPointLabs.NarrationsLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.FYP.Service
{
    public class GenerateVoiceManager : LabCustomizationManager
    {
        public GenerateVoiceManager(string text, int tag, bool isActivated = false)
        {
            this.text = text;
            this.tag = tag;
            this.isActivated = isActivated;
        }
        public override List<Effect> PerformAction(PowerPointSlide slide, int clickNo, string voiceName, bool isSeperateClick = false,
            bool syncAppearance = true)
        {
            if (!syncAppearance)
            {
                return new List<Effect>();
            }
            string name = FYPText.Identifier + FYPText.Underscore + tag.ToString() + FYPText.Underscore + FYPText.AudioIdentifier;
            slide.DeleteShapeWithName(name);
            if (isActivated)
            {             
                return NotesToAudio.EmbedSlideNote(name, text, voiceName, slide, clickNo, isSeperateClick);
            }
            return new List<Effect>();
        }
    }
}
