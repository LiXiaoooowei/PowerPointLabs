using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.FYP.Data
{
    public class BlockItem
    {
        public List<AnimationItem> Items
        {
            get
            {
                return _items;
            }
        }
        public int ClickNo
        {
            get
            {
                return clickNo;
            }
        }
        private List<AnimationItem> _items;
        private int clickNo;

        public BlockItem()
        {
            clickNo = -1;
            _items = new List<AnimationItem>();
        }

        public BlockItem(int clickNo, List<AnimationItem> items)
        {
            this.clickNo = clickNo;
            this._items = items;
        }

        public int GetClickNo()
        {
            return clickNo;
        }

        public List<AnimationItem> GetItems()
        {
            return _items;
        }

        public void InsertItem(AnimationItem item)
        {
            _items.Add(item);
        }
    }
}
