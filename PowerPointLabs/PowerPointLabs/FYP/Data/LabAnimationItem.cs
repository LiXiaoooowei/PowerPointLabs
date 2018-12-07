using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Log;

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
            }
        }
        public bool IsCloudVoice
        {
            get
            {
                return isCloudVoice;
            }
            set
            {
                isCloudVoice = (bool) value;
            }
        }
        public bool IsBuiltInVoice
        {
            get
            {
                return isBuiltInVoice;
            }
            set
            {
                isBuiltInVoice = (bool)value;
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
            }
        }
        public HashSet<Shape> AssociatedShapes
        {
            get
            {
                return associatedShapes;
            }
        }

        private string text;
        private bool isCaption;
        private bool isCloudVoice;
        private bool isBuiltInVoice;
        private bool isCallout;
        private HashSet<Shape> associatedShapes;

        public LabAnimationItem(string text, bool isCaption = false, bool isCloudVoice = false,
            bool isBuiltInVoice = false, bool isCallout = false, HashSet<Shape> shapes = null)
        {
            this.text = text;
            this.isCaption = isCaption;
            this.isCloudVoice = isCloudVoice;
            this.isBuiltInVoice = isBuiltInVoice;
            this.isCallout = isCallout;
            associatedShapes = shapes;
        }

    }
}
