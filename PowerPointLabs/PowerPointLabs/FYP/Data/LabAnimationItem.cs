using System;
using System.Collections.Generic;
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
        public HashSet<Shape> AssociatedShapes
        {
            get
            {
                return associatedShapes;
            }
        }

        public GenerateCalloutManager generateCalloutManager;
        public GenerateCaptionManager generateCaptionManager;
        public GenerateVoiceManager GenerateVoiceManager;
        public int TagNo;

        private string text;
        private bool isCaption;
        private bool isVoice;
        private bool isCallout;
        private HashSet<Shape> associatedShapes;

        public LabAnimationItem(string text, int tagNo, bool isCaption = false, bool isVoice = false,
            bool isCallout = false, HashSet<Shape> shapes = null):base()
        {
            this.text = text;
            TagNo = tagNo;
            this.isCaption = isCaption;
            this.isVoice = isVoice;
            this.isCallout = isCallout;
            associatedShapes = shapes;
            generateCalloutManager = new GenerateCalloutManager(text, tagNo, isCaption);
            generateCaptionManager = new GenerateCaptionManager(text, tagNo, isCaption);
            GenerateVoiceManager = new GenerateVoiceManager(text, tagNo, isVoice);
        }

        public void Execute(PowerPointSlide slide, int clickNo, int seqNo)
        {
            generateCalloutManager.PerformAction(slide, clickNo);
            generateCaptionManager.PerformAction(slide, clickNo);
            GenerateVoiceManager.PerformAction(slide, clickNo, seqNo);
        }

    }
}
