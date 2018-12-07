using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.FYP.Data
{
    public class BlockItemList
    {
        public List<BlockItem> items;

        public BlockItemList(List<BlockItem> items)
        {
            this.items = items;
        }

        public BlockItemList()
        {
            items = new List<BlockItem>();
        }

        public void InsertItem(BlockItem item)
        {
            items.Add(item);
        }

        public List<BlockItem> GetItems()
        {
            return items;
        }
    }
}
