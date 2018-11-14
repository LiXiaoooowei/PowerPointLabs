using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using PowerPointLabs.Tags;

namespace PowerPointLabs.CaptionsLab
{
    public static class NameTagGenerator
    { 

       private static HashSet<int> nameTags = new HashSet<int>();

        public static string GenerateUniqueName()
        {
            int count = 0;
            do
            {
                count++;
            }
            while (nameTags.Contains(count));

            return "PPTLabs Callout " + count.ToString();
        }

        public static void GetTagNo(string note)
        {
            Regex regex = new Regex(@"\[Name\s*:\s*PPTLabs Callout\s*([1-9][0-9]*)\]", RegexOptions.IgnoreCase);
            MatchCollection regexMatches = regex.Matches(note);
            foreach (Match match in regexMatches)
            {
                string value = match.Groups[1].Value.Trim();
                if (!String.IsNullOrEmpty(value))
                {
                    nameTags.Add(Int32.Parse(value));
                }
               
            }
        }
    }
}
