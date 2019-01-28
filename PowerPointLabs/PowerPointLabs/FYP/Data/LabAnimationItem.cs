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
using PowerPointLabs.NarrationsLab.Data;

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
                generateCalloutManager.text = value;
            }
        }
        public string Note
        {
            get
            {
                return note;
            }
            set
            {
                note = value;
                generateCaptionManager.text = value;
                GenerateVoiceManager.text = value;
            }
        }
        public string VoiceLabel
        {
            get
            {
                return voiceLabel;
            }
            set
            {
                voiceLabel = value;
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

        public bool IsTail
        {
            get
            {
                return isTail;
            }
            set
            {
                isTail = (bool)value;
            }
        }

        public GenerateCalloutManager generateCalloutManager;
        public GenerateCaptionManager generateCaptionManager;
        public GenerateVoiceManager GenerateVoiceManager;
        public int TagNo;
        public ObservableCollection<string> AssociatedShapes { get; set; }

        private string text;
        private string note;
        private string voiceLabel;
        private bool isTail;
        private bool isCaption;
        private bool isVoice;
        private bool isCallout;

        public LabAnimationItem(string text, int tagNo, string note = "", bool isCaption = false, bool isVoice = false,
            bool isCallout = false, string voiceLabel = "", bool isTail = false):base()
        {
            this.text = text;
            this.note = note;
            TagNo = tagNo;
            this.isCaption = isCaption;
            this.isVoice = isVoice;
            this.isCallout = isCallout;
            this.voiceLabel = voiceLabel;
            this.isTail = isTail;
            AssociatedShapes = new ObservableCollection<string>();
            generateCalloutManager = new GenerateCalloutManager(text, tagNo, isCallout);
            generateCaptionManager = new GenerateCaptionManager(text, tagNo, isCaption);
            GenerateVoiceManager = new GenerateVoiceManager(text, tagNo, isVoice);
        }

        public void Execute(PowerPointSlide slide, int clickNo, int seqNo, bool isSeperateClick = false, bool syncAppearance = true)
        {         
            bool firstAnimationTriggeredByClick = slide.IsFirstAnimationTriggeredByClick();
            List<Effect> effects = generateCalloutManager.PerformAction(slide, clickNo, isSeperateClick: isSeperateClick, syncAppearance: syncAppearance);
            effects = effects.Concat(generateCaptionManager.PerformAction(slide, clickNo, 
                isSeperateClick: isSeperateClick && !generateCalloutManager.isActivated, syncAppearance: syncAppearance)).ToList();
            effects = effects.Concat(GenerateVoiceManager.PerformAction(slide, clickNo, seqNo, VoiceLabel,
                isSeperateClick: isSeperateClick && !generateCalloutManager.isActivated && !generateCaptionManager.isActivated, syncAppearance: syncAppearance)).ToList();
            if (effects.Count() > 0 && isSeperateClick)
            {
                effects[0].Timing.TriggerType = MsoAnimTriggerType.msoAnimTriggerOnPageClick;
            }
        }

    }
}
