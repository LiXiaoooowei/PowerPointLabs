using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

using Office = Microsoft.Office.Core;

namespace PowerPointLabs.CaptionsLab.CaptionsLabSettings.Storage
{
    class CaptionsLabPresentation : PowerPointPresentation
    {
        private static CaptionsLabPresentation captionsLabPresentation;
        private CaptionsLabPresentation(string path, string filename) : base(path, filename)
        {
        }

        public static CaptionsLabPresentation GetInstance(string path, string filename)
        {
            if (captionsLabPresentation == null)
            {
                captionsLabPresentation = new CaptionsLabPresentation(path, filename);
            }
            return captionsLabPresentation;
        }

        public void AddShape(Shape shape)
        {
            if (!captionsLabPresentation.Opened)
            {
                captionsLabPresentation.Open(withWindow: false, focus: false);
            }
            InitializeSlide();

            PowerPointSlide slide = captionsLabPresentation.Slides[0];
            slide.CopyShapeToSlide(shape);
            captionsLabPresentation.Save();

        }

        public Shape GetShapeWithName(string name)
        {

            if (!captionsLabPresentation.Opened)
            {
                captionsLabPresentation.Open(withWindow: false, focus: false);
            }
            InitializeSlide();

            PowerPointSlide slide = captionsLabPresentation.Slides[0];
            List<Shape> shapes = slide.GetShapeWithName(name);
            Shape shape = shapes.Count > 0 ? shapes[0] : null;
            return shape;

        }

        private bool InitializeSlide()
        {
            if (captionsLabPresentation.Slides.Count == 0)
            {
                captionsLabPresentation.AddSlide();
            }
            return true;
        }
    }
}
