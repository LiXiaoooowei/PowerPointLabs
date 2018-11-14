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

        private Dictionary<string, Callouts> calloutsTable = new Dictionary<string, Callouts>();

        public IntermediateResultTable UpdateNotes(string slideNo, List<Tuple<NameTag, string>> updatedNotes)
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
        public bool IsTableExists(string idx)
        {
            return calloutsTable.ContainsKey(idx);
        }

        private PowerPointCalloutsCache()
        { }
    }
}
