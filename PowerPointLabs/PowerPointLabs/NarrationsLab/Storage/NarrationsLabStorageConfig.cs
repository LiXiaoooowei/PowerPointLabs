using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using PowerPointLabs.NarrationsLab.Data;

namespace PowerPointLabs.NarrationsLab.Storage
{
    public static class NarrationsLabStorageConfig
    {
        private const string defaultNarrationsStorageFolder = "PowerPointLabs Narrations Access Key Storage";
        private const string defaultNarrationsStorageFile = "useraccount.xml";
        private static string defaultApplicationFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        public static string GetAccessKeyStoragePath()
        {
            return Path.Combine(defaultApplicationFolderPath, defaultNarrationsStorageFolder);
        }

        public static string GetAccessKeyFilePath()
        {
            return Path.Combine(defaultApplicationFolderPath, defaultNarrationsStorageFolder, defaultNarrationsStorageFile);
        }

        public static void SaveUserAccount(UserAccount account)
        {
            if (!Directory.Exists(GetAccessKeyStoragePath()))
            {
                Directory.CreateDirectory(GetAccessKeyStoragePath());
            }
            if (!File.Exists(GetAccessKeyFilePath()))
            {
                File.Create(GetAccessKeyFilePath());
            }
        }
    }
}
