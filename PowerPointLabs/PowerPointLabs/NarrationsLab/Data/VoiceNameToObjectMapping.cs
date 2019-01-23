using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.NarrationsLab.Data
{
    public static class VoiceNameToObjectMapping
    {
        public static Dictionary<string, object> VoiceNameToObjectMap =
            new Dictionary<string, object>()
            {
                {
                    Voice.ZiraRUS.ToString(), HumanVoiceList.voices[0]
                },
                {
                    Voice.JessaRUS.ToString(), HumanVoiceList.voices[1]
                },
                {
                    Voice.BenjaminRUS.ToString(), HumanVoiceList.voices[2]
                },
                {
                    Voice.Jessa24kRUS.ToString(), HumanVoiceList.voices[3]
                },
                {
                    Voice.Guy24kRUS.ToString(), HumanVoiceList.voices[4]
                }
            };
    }
}
