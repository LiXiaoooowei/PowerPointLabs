using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        public override void PerformAction(PowerPointSlide slide, int clickNo, int seqNo)
        {
            if (isActivated)
            {
                NotesToAudio.EmbedSlideNote(
                    FYPText.Identifier+FYPText.Underscore+tag.ToString()+FYPText.Underscore+FYPText.AudioIdentifier,
                    text, slide, clickNo, seqNo);
            }
        }
    }
}
