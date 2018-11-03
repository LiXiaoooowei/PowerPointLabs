using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLabs.NarrationsLab.ViewModel
{
    public class GenericEventArgs<T> : EventArgs
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GenericEventArgs{T}" /> class.
        /// </summary>
        /// <param name="eventData">The event data.</param>
        public GenericEventArgs(T eventData, string filepath)
        {
            this.EventData = eventData;
            this.FilePath = filepath;
        }

        public GenericEventArgs(T eventData)
        {
            this.EventData = eventData;
            this.FilePath = null;
        }

        /// <summary>
        /// Gets the event data.
        /// </summary>
        public T EventData { get; private set; }

        public string FilePath { get; private set; }
    }
}
