﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using PowerPointLabs.Models;

namespace PowerPointLabs.FYP.Service
{
    public abstract class LabCustomizationManager
    {
        public string text;
        public int tag;
        public bool isActivated;

        public abstract void PerformAction(PowerPointSlide slide, bool byClick = false);
    }
}
