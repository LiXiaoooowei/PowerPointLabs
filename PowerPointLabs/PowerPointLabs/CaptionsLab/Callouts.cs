using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Tags;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.CaptionsLab
{
    public class Callouts
    {
        private Dictionary<NameTag, string> notes = new Dictionary<NameTag, string>(new NameTag.NameTagEqualityComparator());

        public Callouts(List<Tuple<NameTag, string>> newNotes)
        {
            foreach (Tuple<NameTag, string> note in newNotes)
            {
                notes.Add(note.Item1, note.Item2);
            }
        }

        public IntermediateResultTable UpdateNotes(List<Tuple<NameTag, string>> updatedNotes)
        {
            Dictionary<NameTag, string> notesCopy = new Dictionary<NameTag, string>(new NameTag.NameTagEqualityComparator());
            IntermediateResultTable intermediateResultTable = new IntermediateResultTable();

            foreach (Tuple<NameTag, string> note in updatedNotes)
            {
                if (notes.ContainsKey(note.Item1))
                {
                    Logger.Log("adding to modified note " + note.Item2);
                    ModifyNote(note.Item1, note.Item2);
                    notesCopy[note.Item1] = note.Item2;
                    intermediateResultTable.AddModifiedNote(note.Item1, note.Item2);
                }
                else
                {
                    Logger.Log("inserting new note with key " + note.Item1.Contents);
                    InsertNewNote(note.Item1, note.Item2);
                    notesCopy.Add(note.Item1, note.Item2);
                    intermediateResultTable.AddInsertedNote(note.Item1, note.Item2);
                }
            }

            foreach (KeyValuePair<NameTag, string> note in notes)
            {
                if (!notesCopy.ContainsKey(note.Key))
                {
                    intermediateResultTable.AddDeletedNote(note.Key, note.Value);
                }
            }

            notes = notesCopy;
            intermediateResultTable.AddResultNotes(ToString());
            return intermediateResultTable;
        }

        public override string ToString()
        {
            StringBuilder builder = new StringBuilder();
            foreach (KeyValuePair<NameTag, string> note in notes)
            {
                builder.Append("[Name: " + note.Key.Contents + "] ");
                builder.AppendLine(note.Value + " [AfterClick]");
            }
            return builder.ToString();
        }

        public bool IsEmpty()
        {
            return notes.Count == 0;
        }

        private bool InsertNewNote(NameTag tag, string note)
        {
            if (notes.ContainsKey(tag))
            {
                return false;
            }
            notes.Add(tag, note);
            return true;
        }

        private bool DeleteNote(NameTag tag)
        {
            if (notes.ContainsKey(tag))
            {
                notes.Remove(tag);
                return true;
            }
            return false;
        }

        private bool ModifyNote(NameTag tag, string modifiedNote)
        {
            if (!notes.ContainsKey(tag))
            {
                return false;
            }
            notes[tag] = modifiedNote;
            return true;
        }
    }
}
