using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.FYP.Service
{
    class LabAnimationItemIdentifierManager
    {
        private static HashSet<int> nameTags = new HashSet<int>();

        public static int GenerateUniqueNumber()
        {
            int count = 0;
            do
            {
                count++;
            }
            while (nameTags.Contains(count));

            nameTags.Add(count);
            return count;
        }

        public static int GetTagNo(string name)
        {
            Regex regex = new Regex(FYPText.Identifier+
                FYPText.Underscore+@"([1-9][0-9]*)"+FYPText.Underscore
                +"("+FYPText.CalloutIdentifier+"|"
                +FYPText.CaptionIdentifier+"|"
                +FYPText.AudioIdentifier+")", RegexOptions.IgnoreCase);
            Match match = regex.Match(name);
            int value = -1;
            if (match.Success)
            {
                try
                {
                    value = int.Parse(match.Groups[1].Value.Trim());
                    nameTags.Add(value);
                }
                catch
                { }
            }
            return value;
        }

        public static string GetTagFunction(string name)
        {
            Regex regex = new Regex(FYPText.Identifier +
               FYPText.Underscore + @"[1-9][0-9]*" +FYPText.Underscore
               + "(" + FYPText.CalloutIdentifier + "|"
               + FYPText.CaptionIdentifier + "|"
               + FYPText.AudioIdentifier + ")", RegexOptions.IgnoreCase);
            Match match = regex.Match(name);
            string value = "";
            if (match.Success)
            {
                value = match.Groups[1].Value.Trim();
            }
            return value;
        }

        public static void EmptyTagsCollection()
        {
            nameTags.Clear();
        }
    }
}
