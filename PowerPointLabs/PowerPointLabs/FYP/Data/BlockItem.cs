using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.FYP.Data
{
    public class BlockItem
    {
        public ObservableCollection<AnimationItem> Items
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
        private ObservableCollection<AnimationItem> _items;
        private int clickNo;

        public BlockItem()
        {
            clickNo = -1;
            _items = new ObservableCollection<AnimationItem>();
        }

        public BlockItem(int clickNo, ObservableCollection<AnimationItem> items)
        {
            this.clickNo = clickNo;
            this._items = items;
        }

        public int GetClickNo()
        {
            return clickNo;
        }

        public void InsertItem(AnimationItem item)
        {
            _items.Add(item);
        }

        public bool IsEmpty()
        {
            return _items.Count == 0;
        }

        public void SetClickNo(int no)
        {
            clickNo = no;
        }
    }
}
