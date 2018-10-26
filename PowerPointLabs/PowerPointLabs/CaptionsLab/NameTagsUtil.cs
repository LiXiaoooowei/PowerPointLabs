using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using PowerPointLabs.Tags;

namespace PowerPointLabs.CaptionsLab
{
    public static class NameTagsUtil
    { 

       private static HashSet<NameTag> nameTags = new HashSet<NameTag>();

        private static int count = 0;

        public static string GenerateUniqueName()
        {
            NameTag tag;

            do
            {
                tag = new NameTag(0, 0, "PPTLabs Callout " + (++count).ToString());
            }
            while (nameTags.Contains(tag));

            return tag.Contents;
        }

        public static void InitializeCount(int cnt)
        {
            count = cnt;
        }

        public static int GetTagNo(string note)
        {
            int no = 0;
            Regex regex = new Regex(@"\[Name\s*:\s*PPTLabs Callout\s*([1-9][0-9]*)\]", RegexOptions.IgnoreCase);
            MatchCollection regexMatches = regex.Matches(note);
            foreach (Match match in regexMatches)
            {
                string value = match.Groups[1].Value.Trim();
                int num = String.IsNullOrEmpty(value) ? -1 : Int32.Parse(value);
                if (no < num)
                {
                    no = num;
                }
            }
            return no;
        }
    }
}
