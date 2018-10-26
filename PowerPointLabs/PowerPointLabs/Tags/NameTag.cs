using System;
using System.Collections.Generic;
using System.Speech.Synthesis;

namespace PowerPointLabs.Tags
{
    public class NameTag: Tag
    {
        public NameTag(int start, int end, string contents)
        {
            Start = start;
            End = end;
            Contents = contents;
        }
        public override bool Apply(PromptBuilder builder)
        {
            return true;
        }

        public override string PrettyPrint()
        {
            return "";
        }

        public class NameTagEqualityComparator : IEqualityComparer<NameTag>
        {
            public bool Equals(NameTag x, NameTag y)
            {
                return x.Contents.Equals(y.Contents);
            }

            public int GetHashCode(NameTag obj)
            {
                return obj.Contents.GetHashCode();
            }
        }
    }
}
