using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Media;
using System.Threading;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Models;
using PowerPointLabs.NarrationsLab;
using PowerPointLabs.NarrationsLab.Data;
using PowerPointLabs.NarrationsLab.ViewModel;
using PowerPointLabs.TextCollection;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.ActionFramework.NarrationsLab
{
    [ExportActionRibbonId("TestHumanNarrationsButton")]
    class TestHumanNarrationsActionHandler : ActionHandler
    {
#pragma warning disable 0618
        protected override void ExecuteAction(string ribbonId)
        {
            //TODO: This needs to improved to stop using global variables
            this.StartNewUndoEntry();

            PowerPointSlide currentSlide = this.GetCurrentSlide();

            // If there are text in notes page for any of the selected slides 
            if (this.GetCurrentPresentation().SelectedSlides.Any(slide => slide.NotesPageText.Trim() != ""))
            {
                NotesToAudio.IsRemoveAudioEnabled = true;
                this.GetRibbonUi().RefreshRibbonControl("RemoveNarrationsButton");
            }
           
            List<string[]> audioList = new List<string[]>();

            List<PowerPointSlide> slides = PowerPointCurrentPresentationInfo.SelectedSlides.ToList();

            int numberOfSlides = slides.Count;
            for (int currentSlideIndex = 0; currentSlideIndex < numberOfSlides; currentSlideIndex++)
            {
                PowerPointSlide slide = slides[currentSlideIndex];
                string accessToken;
                Authentication auth = new Authentication("https://westus.api.cognitive.microsoft.com/sts/v1.0/issueToken", "9b3f23b8a9b14b32b40fee2ba141b5ac");

                try
                {
                    accessToken = auth.GetAccessToken();
                    Console.WriteLine("Token: {0}\n", accessToken);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Failed authentication.");
                    Console.WriteLine(ex.ToString());
                    Console.WriteLine(ex.Message);
                    return;
                }

                string requestUri = "https://westus.tts.speech.microsoft.com/cognitiveservices/v1";
                var cortana = new Synthesize();

                cortana.OnAudioAvailable += SaveAudioToWaveFile;
                cortana.OnError += ErrorHandler;

                // Reuse Synthesize object to minimize latency
                cortana.Speak(CancellationToken.None, new Synthesize.InputOptions()
                {
                    RequestUri = new Uri(requestUri),
                    // Text to be spoken.
                    Text = slide.NotesPageText,
                    VoiceType = Gender.Female,
                    // Refer to the documentation for complete list of supported locales.
                    Locale = "en-US",
                    // You can also customize the output voice. Refer to the documentation to view the different
                    // voices that the TTS service can output.
                    // VoiceName = "Microsoft Server Speech Text to Speech Voice (en-US, Jessa24KRUS)",
                    VoiceName = "Microsoft Server Speech Text to Speech Voice (en-US, Guy24KRUS)",
                    // VoiceName = "Microsoft Server Speech Text to Speech Voice (en-US, ZiraRUS)",

                    // Service can return audio in different output format.
                    OutputFormat = AudioOutputFormat.Riff24Khz16BitMonoPcm,
                    AuthorizationToken = "Bearer " + accessToken,
                }, "").Wait();

                Shape audioShape = InsertAudioFileOnSlide(slide, @"C:\Users\xiaov\Desktop\test.wav");
            }
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
            SaveStreamToFile(@"C:\Users\xiaov\Desktop\test.wav", args.EventData);
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

        private static Shape InsertAudioFileOnSlide(PowerPointSlide slide, string fileName)
        {
            float slideWidth = PowerPointPresentation.Current.SlideWidth;

            Shape audioShape = slide.Shapes.AddMediaObject2(fileName, MsoTriState.msoFalse, MsoTriState.msoTrue, slideWidth + 20);
            slide.RemoveAnimationsForShape(audioShape);

            return audioShape;
        }
    }
}
