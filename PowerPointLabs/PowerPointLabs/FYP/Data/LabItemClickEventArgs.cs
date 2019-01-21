using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.FYP.Data
{
    public class LabItemClickEventArgs: EventArgs
    {
        private LabAnimationItem item;
        public LabItemClickEventArgs(LabAnimationItem item)
        {
            this.item = item;
        }
        public LabAnimationItem GetData()
        {
            return item;
        }
    }
}
