using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Speech.Synthesis;
using System.Text;
using System.Threading;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.NarrationsLab;
using PowerPointLabs.NarrationsLab.Data;
using PowerPointLabs.NarrationsLab.ViewModel;

namespace PowerPointLabs.SpeechEngine
{
    static class TextToSpeech
    {
        public static String DefaultVoiceName;
        public static HumanVoice humanVoice;

        public static IEnumerable<string> GetVoices()
        {
            using (SpeechSynthesizer synthesizer = new SpeechSynthesizer())
            {
                System.Collections.ObjectModel.ReadOnlyCollection<InstalledVoice> installedVoices = synthesizer.GetInstalledVoices();
                IEnumerable<InstalledVoice> voices = installedVoices.Where(voice => voice.Enabled);
                return voices.Select(voice => voice.VoiceInfo.Name);
            }
        }

        public static void SaveStringToWaveFiles(string notesText, string folderPath, string fileNameFormat)
        {
            TaggedText taggedNotes = new TaggedText(notesText);
            List<String> stringsToSave = taggedNotes.SplitByClicks();
            //MD5 md5 = MD5.Create();

            for (int i = 0; i < stringsToSave.Count; i++)
            {
                String textToSave = stringsToSave[i];
                String baseFileName = String.Format(fileNameFormat, i + 1);

                // The first item will autoplay; everything else is triggered by a click.
                String fileName = i > 0 ? baseFileName + " (OnClick)" : baseFileName;

                String filePath = folderPath + "\\" + fileName + ".wav";
                if (!NotesToAudio.IsHumanVoiceSelected)
                {
                    SaveStringToWaveFile(textToSave, filePath);
                }
                else
                {
                    SaveStringToWaveFileWithHumanVoice(textToSave, filePath);
                }
            }
        }

        public static void SaveStringToWaveFile(String textToSave, String filePath)
        {
            PromptBuilder builder = GetPromptForText(textToSave);
            PromptToAudio.SaveAsWav(builder, filePath);
        }

        public static void SpeakString(String textToSpeak)
        {
            if (String.IsNullOrWhiteSpace(textToSpeak))
            {
                return;
            }

            PromptBuilder builder = GetPromptForText(textToSpeak);
            PromptToAudio.Speak(builder);
        }


        private static void SaveStringToWaveFileWithHumanVoice(string textToSave, string filePath)
        {
            string accessToken;
            string textToSpeak = GetHumanSpeakNotesForText(textToSave);
            Authentication auth = new Authentication(UserAccount.GetInstance().GetEndpoint(), UserAccount.GetInstance().GetKey());

            try
            {
                accessToken = auth.GetAccessToken();
                Logger.Log("Token: " + accessToken);
            }
            catch (Exception ex)
            {
                Logger.Log("Failed authentication.");
                Logger.Log(ex.ToString());
                Logger.Log(ex.Message);
                return;
            }

            string requestUri = UserAccount.GetInstance().GetUri();
            var cortana = new Synthesize();

            cortana.OnAudioAvailable += SaveAudioToWaveFile;
            cortana.OnError += ErrorHandler;

            // Reuse Synthesize object to minimize latency
            cortana.Speak(CancellationToken.None, new Synthesize.InputOptions()
            {
                RequestUri = new Uri(requestUri),
                Text = textToSpeak,
                VoiceType = humanVoice.voiceType,
                Locale = humanVoice.Locale,
                VoiceName = humanVoice.voiceName,
                // Service can return audio in different output format.
                OutputFormat = AudioOutputFormat.Riff24Khz16BitMonoPcm,
                AuthorizationToken = "Bearer " + accessToken,
            }, filePath).Wait();

        }

        private static string GetHumanSpeakNotesForText(string textToSave)
        {
            TaggedText taggedText = new TaggedText(textToSave);
            string strToSpeak = taggedText.ToPrettyString();
            return strToSpeak;
        }

        private static PromptBuilder GetPromptForText(string textToConvert)
        {
            TaggedText taggedText = new TaggedText(textToConvert);
            PromptBuilder builder = taggedText.ToPromptBuilder(DefaultVoiceName);
            return builder;
        }

        private static byte[] ReadFully(Stream input)
        {
            byte[] buffer = new byte[16 * 1024];

            using (MemoryStream ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return ms.ToArray();
            }
        }
        private static void SaveStreamToFile(string fileFullPath, Stream stream)
        {
            try
            {
                byte[] bytesInStream = ReadFully(stream);
                using (FileStream fs = File.Create(fileFullPath))
                {
                    Console.WriteLine("file created");
                    fs.Write(bytesInStream, 0, bytesInStream.Length);
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }
        /// <summary>
        /// This method is called once the audio returned from the service.
        /// It will then attempt to play that audio file.
        /// Note that the playback will fail if the output audio format is not pcm encoded.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="args">The <see cref="GenericEventArgs{Stream}"/> instance containing the event data.</param>
        private static void SaveAudioToWaveFile(object sender, GenericEventArgs<Stream> args)
        {
            Console.WriteLine(args.EventData);
            SaveStreamToFile(args.FilePath, args.EventData);
            Console.WriteLine("saving to wav");
            // For SoundPlayer to be able to play the wav file, it has to be encoded in PCM.
            // Use output audio format AudioOutputFormat.Riff16Khz16BitMonoPcm to do that.
            //   SoundPlayer player = new SoundPlayer(args.EventData);
            //  player.PlaySync();
            args.EventData.Dispose();
        }

        /// <summary>
        /// Handler an error when a TTS request failed.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="GenericEventArgs{Exception}"/> instance containing the event data.</param>
        private static void ErrorHandler(object sender, GenericEventArgs<Exception> e)
        {
            Console.WriteLine("Unable to complete the TTS request: [{0}]", e.ToString());
        }
    }
}