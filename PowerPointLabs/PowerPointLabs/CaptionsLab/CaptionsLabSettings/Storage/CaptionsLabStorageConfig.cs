using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ShortcutsLab;
using PowerPointLabs.Utils;

namespace PowerPointLabs.CaptionsLab.CaptionsLabSettings.Storage
{
    public static class CaptionsLabStorageConfig
    {
        private const string defaultCalloutStorageFolder = "PowerPointLabs Callout Storage";
        private const string defaultCalloutImageFolder = "Callout Images";
        private static string defaultApplicationFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        public static string GetCalloutImageStoragePath()
        {
            return Path.Combine(defaultApplicationFolderPath, defaultCalloutStorageFolder, defaultCalloutImageFolder);
        }

        public static string GetCalloutPptxStoragePath()
        {
            return Path.Combine(defaultApplicationFolderPath, defaultCalloutStorageFolder);
        }

        public static void SaveSelectedShapeConfig(Shape shape)
        {
            if (!Directory.Exists(GetCalloutImageStoragePath()))
            {
                Directory.CreateDirectory(GetCalloutImageStoragePath());
            }
            string shapeFullName = GetCalloutImageStoragePath() + @"\" + shape.Name + ".png";
            GraphicsUtil.ExportShape(shape, shapeFullName);
            CaptionsLabPresentation pres = CaptionsLabPresentation.GetInstance(GetCalloutPptxStoragePath(), "new");
            pres.AddShape(shape);
        }
    }
}
