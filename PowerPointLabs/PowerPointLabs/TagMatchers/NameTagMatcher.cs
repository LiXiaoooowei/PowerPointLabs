using System.Collections.Generic;
using System.Text.RegularExpressions;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Tags;

namespace PowerPointLabs.TagMatchers
{
    public class NameTagMatcher: ITagMatcher
    {
        public Regex Regex { get { return new Regex(@"\[Name\s*:(.*)\]", RegexOptions.IgnoreCase); } }

        public List<ITag> Matches(string text)
        {
            List<ITag> foundMatches = new List<ITag>();

            MatchCollection regexMatches = Regex.Matches(text);
            foreach (Match match in regexMatches)
            {
                int matchStart = match.Index;
                int matchEnd = match.Index + match.Length - 1; // 0-based indices.
                NameTag tag = new NameTag(matchStart, matchEnd, match.Groups[1].Value.Trim());
                foundMatches.Add(tag);
            }

            return foundMatches;
        }

        public List<NameTag> NameTagMatches(string text)
        {
            List<NameTag> foundMatches = new List<NameTag>();

            MatchCollection regexMatches = Regex.Matches(text);
            foreach (Match match in regexMatches)
            {
                int matchStart = match.Index;
                int matchEnd = match.Index + match.Length - 1; // 0-based indices.
                NameTag tag = new NameTag(matchStart, matchEnd, match.Groups[1].Value.Trim());
                foundMatches.Add(tag);
            }

            return foundMatches;
        }

    }
}
