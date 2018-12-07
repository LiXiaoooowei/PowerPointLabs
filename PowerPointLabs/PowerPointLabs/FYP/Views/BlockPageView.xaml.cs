using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.FYP.Data;
using PowerPointLabs.Models;

namespace PowerPointLabs.FYP.Views
{
#pragma warning disable 0618
    /// <summary>
    /// Interaction logic for BlockPageView.xaml
    /// </summary>
    public partial class BlockPageView : UserControl
    {
        public List<BlockItem> BlockItems
        {
            get
            {
                return blockItemList.GetItems();
            }
            
        }
        private BlockItemList blockItemList;

        public BlockPageView()
        {
            InitializeComponent();
            blockItemList = InitializeBlockItemList();
            Globals.ThisAddIn.Application.SlideSelectionChanged += Handle;
            listView.ItemsSource = BlockItems;
        }

        private BlockItemList InitializeBlockItemList()
        {
            IEnumerable<Effect> effects = PowerPointCurrentPresentationInfo.CurrentSlide.TimeLine.MainSequence.Cast<Effect>();
            BlockItemList list = new BlockItemList();
            List<AnimationItem> items = new List<AnimationItem>();
            int clickNo = 0;
            for (int i = 0; i < effects.Count(); i++)
            {              
                Effect effect = effects.ElementAt(i);
                if (effect.Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerOnPageClick)
                {
                    if (items.Count > 0)
                    {
                        list.InsertItem(new BlockItem(clickNo, items.GetRange(0, items.Count)));
                    }
                    items.Clear();
                    clickNo++;
                }
                items.Add(new CustomAnimationItem(effect));
            }
            if (items.Count > 0)
            {
                list.InsertItem(new BlockItem(clickNo, items));
            }
            return list;
        }

        private void Handle(Microsoft.Office.Interop.PowerPoint.SlideRange sldRange)
        {
            if (PowerPointCurrentPresentationInfo.CurrentSlide != null)
            {
                blockItemList = InitializeBlockItemList();
                listView.ItemsSource = null;
                listView.ItemsSource = BlockItems;
            }
        }
    }
}
