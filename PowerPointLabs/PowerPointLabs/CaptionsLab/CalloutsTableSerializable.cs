using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using PowerPointLabs.Models;

namespace PowerPointLabs.CaptionsLab
{
    public class CalloutsTableSerializable
    {
        public class CalloutsListSerializable
        {
            public int slideNo;
            public CalloutSerializable callout;
        }

        public class CalloutSerializable
        {
            public int shapeCount;
            public List<string> notesInverted;
            public List<IntPair> calloutsInverted;
            public Callouts ToCallouts()
            {
                Dictionary<int, int> pairs = new Dictionary<int, int>();
                foreach (IntPair pair in calloutsInverted)
                {
                    pairs.Add(pair.stmtNo, pair.shapeNo);
                }
                return new Callouts(shapeCount, pairs, notesInverted);
            }
        }

        public class IntPair
        {
            public int stmtNo;
            public int shapeNo;
        }

        public List<CalloutsListSerializable> list;

        public Dictionary<int, Callouts> ToCalloutsTable()
        {
            Dictionary<int, Callouts> calloutsTable = new Dictionary<int, Callouts>();
            foreach (CalloutsListSerializable item in list)
            {
                int slideNo = item.slideNo;
                Callouts callouts = item.callout.ToCallouts();
                calloutsTable.Add(slideNo, callouts);
            }
            return calloutsTable;
        }
    }
}
