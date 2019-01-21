using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.FYP.Service;
using PowerPointLabs.Models;

namespace PowerPointLabs.FYP.Data
{
    public class LabAnimationItem: AnimationItem
    {
        public string Text
        {
            get
            {
                return text;
            }
            set
            {
                text = value;
                generateCaptionManager.text = value;
                generateCalloutManager.text = value;
                GenerateVoiceManager.text = value;
            }
        }
        public bool IsCaption
        {
            get
            {
                return isCaption;
            }
            set
            {
                isCaption = (bool)value;
                generateCaptionManager.isActivated = (bool)value;
            }
        }
        public bool IsVoice
        {
            get
            {
                return isVoice;
            }
            set
            {
                isVoice = (bool) value;
                GenerateVoiceManager.isActivated = (bool)value;
            }
        }

        public bool IsCallout
        {
            get
            {
                return isCallout;
            }
            set
            {
                isCallout = (bool)value;
                generateCalloutManager.isActivated = (bool)value;
            }
        }

        public GenerateCalloutManager generateCalloutManager;
        public GenerateCaptionManager generateCaptionManager;
        public GenerateVoiceManager GenerateVoiceManager;
        public int TagNo;
        public ObservableCollection<string> AssociatedShapes { get; set; }

        private string text;
        private bool isCaption;
        private bool isVoice;
        private bool isCallout;

        public LabAnimationItem(string text, int tagNo, bool isCaption = false, bool isVoice = false,
            bool isCallout = false):base()
        {
            this.text = text;
            TagNo = tagNo;
            this.isCaption = isCaption;
            this.isVoice = isVoice;
            this.isCallout = isCallout;
            AssociatedShapes = new ObservableCollection<string>();
            generateCalloutManager = new GenerateCalloutManager(text, tagNo, isCallout);
            generateCaptionManager = new GenerateCaptionManager(text, tagNo, isCaption);
            GenerateVoiceManager = new GenerateVoiceManager(text, tagNo, isVoice);
        }

        public void Execute(PowerPointSlide slide, int clickNo, int seqNo, bool isSeperateClick = false)
        {         
            bool firstAnimationTriggeredByClick = slide.IsFirstAnimationTriggeredByClick();
            List<Effect> effects = generateCalloutManager.PerformAction(slide, clickNo, isSeperateClick: isSeperateClick);
            effects = effects.Concat(generateCaptionManager.PerformAction(slide, clickNo, 
                isSeperateClick: isSeperateClick && !generateCalloutManager.isActivated)).ToList();
            effects = effects.Concat(GenerateVoiceManager.PerformAction(slide, clickNo, seqNo, 
                isSeperateClick: isSeperateClick && !generateCalloutManager.isActivated && !generateCaptionManager.isActivated)).ToList();
            IEnumerable<Effect> allEffects = slide.TimeLine.MainSequence.Cast<Effect>();
            if (clickNo <= 0 && isSeperateClick && !firstAnimationTriggeredByClick)
            {
                allEffects.ElementAt(effects.Count()).Timing.TriggerType = MsoAnimTriggerType.msoAnimTriggerOnPageClick;
            }
            else if (isSeperateClick && effects.Count() > 0)
            {
                effects[0].Timing.TriggerType = MsoAnimTriggerType.msoAnimTriggerOnPageClick;
            }
        }

    }
}
