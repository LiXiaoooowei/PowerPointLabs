using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using PowerPointLabs.Models;
using PowerPointLabs.Tags;

namespace PowerPointLabs.CaptionsLab
{
    public sealed class PowerPointCalloutsCache
    {
        public static PowerPointCalloutsCache Instance
        {
            get
            {
                return powerPointCalloutsCache.Value;
            }
        }

        private static Lazy<PowerPointCalloutsCache> powerPointCalloutsCache 
            = new Lazy<PowerPointCalloutsCache>(() => new PowerPointCalloutsCache());

        private Dictionary<int, Callouts> calloutsTable = new Dictionary<int, Callouts>();
        private int slideCnt = 0;
        public IntermediateResultTable UpdateNotes(int slideNo, List<Tuple<NameTag, string>> updatedNotes)
        {
            IntermediateResultTable context = new IntermediateResultTable();
            IEnumerable<NameTag> notes = from note in updatedNotes select note.Item1;
            if (IsTableExists(slideNo))
            {
                context = calloutsTable[slideNo].UpdateNotes(updatedNotes);
                if (calloutsTable[slideNo].IsEmpty())
                {
                    calloutsTable.Remove(slideNo);
                }
            }
            else
            {
                calloutsTable.Add(slideNo, new Callouts(updatedNotes));
                context.AddInsertedNote(updatedNotes);
                context.AddResultNotes(calloutsTable[slideNo].ToString());
            }
            context.SetNotes(notes);
            return context;
        }
        public bool IsTableExists(int idx)
        {
            return calloutsTable.ContainsKey(idx);
        }

        public int CreateNewTableEntry()
        {
            return ++slideCnt; 
        }

        public void InitializeSlideCount(int cnt)
        {
            slideCnt = cnt;
        }

        private PowerPointCalloutsCache()
        { }
    }
}
