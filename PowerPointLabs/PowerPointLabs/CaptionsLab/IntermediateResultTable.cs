using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using PowerPointLabs.Tags;

namespace PowerPointLabs.CaptionsLab
{
    public class IntermediateResultTable
    {
        List<Tuple<NameTag, string>> notesInserted = new List<Tuple<NameTag, string>>();
        List<Tuple<NameTag, string>> notesDeleted = new List<Tuple<NameTag, string>>();
        List<Tuple<NameTag, string>> notesModified = new List<Tuple<NameTag, string>>();
        string resultNotes = "";

        public void AddInsertedNote(NameTag tag, string note)
        {
            notesInserted.Add(new Tuple<NameTag, string>(tag, note));
        }

        public void AddInsertedNote(List<Tuple<NameTag, string>> lists)
        {
            notesInserted = lists;
        }

        public void AddDeletedNote(NameTag tag, string note)
        {
            notesDeleted.Add(new Tuple<NameTag, string>(tag, note));
        }

        public void AddModifiedNote(NameTag tag, string note)
        {
            notesModified.Add(new Tuple<NameTag, string>(tag, note));
        }

        public void AddResultNotes(string note)
        {
            resultNotes = note;
        }

        public string GetResultNotes()
        {
            return resultNotes;
        }

        public List<Tuple<NameTag, string>> GetDeletedNotes()
        {
            return notesDeleted;
        }

        public List<Tuple<NameTag, string>> GetInsertedNotes()
        {
            return notesInserted;
        }

        public List<Tuple<NameTag, string>> GetModifiedNotes()
        {
            return notesModified;
        }
    }
}
