using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.Models
{
    public class Callouts
    {
        // map from stmt idx to callout shape idx
        private Dictionary<int, int> calloutsInverted = new Dictionary<int, int>();
        // map from callout idx to note string idx
     //   private Dictionary<int, int> calloutsInverted = new Dictionary<int, int>();
        // map from note string to note string idx
   //     private Dictionary<string, int> notes = new Dictionary<string, int>();
        // map from stmt idx to note string
        private List<string> notesInverted = new List<string>();

        public Callouts()
        {
        }

        public int InsertCallout(string note, int stmtIdx)
        {
            notesInverted.Insert(stmtIdx, note);
            int calloutIdx = calloutsInverted.Count + 1;
            calloutsInverted.Add(stmtIdx, calloutIdx);
            return calloutIdx;
        }

        public int GetCalloutNoFromStmtNo(int stmtNo)
        {
            return calloutsInverted.ContainsKey(stmtNo) ? calloutsInverted[stmtNo] : -1;
        }
    }
}
