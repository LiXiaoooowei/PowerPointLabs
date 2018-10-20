using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;

namespace PowerPointLabs.CaptionsLab
{
    public static class StorageUtil
    {
        //https://stackoverflow.com/questions/6115721/how-to-save-restore-serializable-object-to-from-file
        public static void WriteToXMLFile(string filename, CalloutsTableSerializable objectToWrite, bool append = false)
        {
            TextWriter writer = null;
            try
            {
                var serializer = new XmlSerializer(typeof(CalloutsTableSerializable));
                writer = new StreamWriter(filename, append);
                serializer.Serialize(writer, objectToWrite);
            }
            catch (Exception e)
            {
                Logger.Log(e.Message);
            }
            finally
            {
                if (writer != null)
                {
                    writer.Close();
                }
            }
        }

        //https://stackoverflow.com/questions/6115721/how-to-save-restore-serializable-object-to-from-file
        public static T ReadFromXmlFile<T>(string filename) where T : new()
        {
            TextReader reader = null;
            try
            {
                var serializer = new XmlSerializer(typeof(T));
                reader = new StreamReader(filename);
                return (T)serializer.Deserialize(reader);
            }
            catch (Exception)
            {
                return default(T);
            }
            finally
            {
                if (reader != null)
                {
                    reader.Close();
                }
            }
        }
    }
}
