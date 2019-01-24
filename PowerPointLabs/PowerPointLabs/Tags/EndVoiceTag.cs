﻿using System.Speech.Synthesis;

using PowerPointLabs.SpeechEngine;

namespace PowerPointLabs.Tags
{
    public class EndVoiceTag : Tag
    {
        public EndVoiceTag(int start, int end, string contents)
        {
            Start = start;
            End = end;
            Contents = contents;
        }

        public override bool Apply(PromptBuilder builder)
        {
            builder.EndVoice();
            builder.StartVoice(TextToSpeech.DefaultVoiceName);
            return true;
        }

        public override string PrettyPrint()
        {
            return "";
        }
    }
}
