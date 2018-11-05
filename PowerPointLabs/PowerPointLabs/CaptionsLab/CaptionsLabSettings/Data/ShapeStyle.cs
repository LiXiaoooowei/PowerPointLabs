using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLabs.CaptionsLab.CaptionsLabSettings.Data
{
    public class ShapeStyle
    {
        public ShapeStyle(string path)
        {
            Source = path;
        }

        public string Source { get; }

        public override string ToString() => Source;
    }
}
