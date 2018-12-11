using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.FYP.Data
{
    public class BlockItemList
    {

        public ObservableCollection<BlockItem> items;

        public BlockItemList(ObservableCollection<BlockItem> items)
        {
            this.items = items;
        }

        public BlockItemList()
        {
            items = new ObservableCollection<BlockItem>();
        }

        public void InsertItem(BlockItem item)
        {
            items.Add(item);
        }

    }
}
