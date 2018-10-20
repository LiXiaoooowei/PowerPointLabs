using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using PowerPointLabs.ActionFramework.Common.Log;
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
        int shapeCount = 0;
        public Callouts()
        {
        }

        public int InsertCallout(string note, int stmtIdx)
        {           
            for (int i = notesInverted.Count - 1; i >= stmtIdx; i--)
            {
                int calloutShapeIdx = calloutsInverted[i];
                calloutsInverted.Remove(i);
                calloutsInverted.Add(i + 1, calloutShapeIdx);
            }
            notesInverted.Insert(stmtIdx, note);
            int calloutIdx = (++shapeCount);
            calloutsInverted.Add(stmtIdx, calloutIdx);
            return calloutIdx;
        }

        public int InsertCallout(string note)
        {
            int stmtIdx = notesInverted.Count;
            notesInverted.Add(note);
            int calloutIdx = (++shapeCount);
            calloutsInverted.Add(stmtIdx, calloutIdx);
            return calloutIdx;
        }

        public int DeleteCallout(int stmtIdx)
        {
            int calloutIdx = calloutsInverted[stmtIdx];
            for (int i = stmtIdx + 1; i < notesInverted.Count; i++)
            {
                int calloutShapeIdx = calloutsInverted[i];
                calloutsInverted.Remove(i - 1);
                calloutsInverted.Add(i - 1, calloutShapeIdx);
            }
            calloutsInverted.Remove(notesInverted.Count - 1);
            notesInverted.RemoveAt(stmtIdx);
            return calloutIdx;
        }

        public int UpdateCallout(string note, int stmtIdx)
        {
            int calloutIdx = calloutsInverted[stmtIdx];
            notesInverted[stmtIdx] = note;
            return calloutIdx;
        }

        public int GetNotesInvertedCount()
        {
            return notesInverted.Count;
        }

        public int GetCalloutIdxFromStmtNo(int stmtNo)
        {
            return calloutsInverted[stmtNo];
        }

        public string NotesToString()
        {
            StringBuilder builder = new StringBuilder();
            foreach (string s in notesInverted)
            {
                string _s = s.Replace("[i]", "");
                builder.Append(_s + "[AfterClick]" + " ");               
            }
            return builder.ToString();
        }
    }
}
