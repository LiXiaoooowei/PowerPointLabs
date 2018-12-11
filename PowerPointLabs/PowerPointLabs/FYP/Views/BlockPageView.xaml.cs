using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Media;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.FYP.Data;
using PowerPointLabs.FYP.Service;
using PowerPointLabs.Models;

namespace PowerPointLabs.FYP.Views
{
#pragma warning disable 0618
    /// <summary>
    /// Interaction logic for BlockPageView.xaml
    /// </summary>
    public partial class BlockPageView : UserControl
    {
        public ObservableCollection<BlockItem> BlockItems
        {
            get
            {
                return blockItemList.items;
            }
        }
        private ListView draggedListView = null;
        private int draggedListViewIndex = -1;
        private BlockItemList blockItemList;
        private ListViewDragDropManager<BlockItem> blockDragManager;


        public BlockPageView()
        {
            if (LicenseManager.UsageMode == LicenseUsageMode.Runtime)
            {
                InitializeComponent();
                blockItemList = InitializeBlockItemList();
                Globals.ThisAddIn.Application.SlideSelectionChanged += Handle;
                listView.ItemsSource = BlockItems;
            }
        }

        public void HandleButtonClick()
        {
            Logger.Log("clicked");
            PowerPointSlide slide = PowerPointCurrentPresentationInfo.CurrentSlide;
            IEnumerable<Effect> effects = slide.TimeLine.MainSequence.Cast<Effect>();
            IEnumerable<Shape> shapes = slide.Shapes.Cast<Shape>();
            slide.RemoveAnimationsForShapes(shapes.ToList());
            
            ObservableCollection<BlockItem> animationItems = listView.ItemsSource as ObservableCollection<BlockItem>;
            
            for (int i = 0; i < animationItems.Count(); i++)
            {
                BlockItem blockItem = animationItems.ElementAt(i);
                for (int j = 0; j < blockItem.Items.Count; j++)
                {
                    CustomAnimationItem item = blockItem.Items.ElementAt(j) as CustomAnimationItem;
                    if (j == 0)
                    {
                        Effect effect = slide.TimeLine.MainSequence.AddEffect(item.GetShape(), item.GetEffectType(), 
                            item.GetEffectLevel(), MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                        effect.Exit = item.GetExit();
                    }
                    else
                    {
                        Effect effect = slide.TimeLine.MainSequence.AddEffect(item.GetShape(), item.GetEffectType(), 
                            item.GetEffectLevel(), MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                        try
                        {
                            effect.Exit = item.GetExit();
                        }
                        catch
                        { }
                    }
                }
            }
        }

        private BlockItemList InitializeBlockItemList()
        {
            IEnumerable<Effect> effects = PowerPointCurrentPresentationInfo.CurrentSlide.TimeLine.MainSequence.Cast<Effect>();
            BlockItemList list = new BlockItemList();
            ObservableCollection<AnimationItem> items = new ObservableCollection<AnimationItem>();
            int clickNo = 0;
            for (int i = 0; i < effects.Count(); i++)
            {              
                Effect effect = effects.ElementAt(i);
                if (effect.Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerOnPageClick)
                {
                    if (items.Count > 0)
                    {
                        list.InsertItem(new BlockItem(clickNo, new ObservableCollection<AnimationItem>(items)));
                    }
                    items.Clear();
                    clickNo++;
                }
                items.Add(new CustomAnimationItem(effect.Shape,
                           effect.EffectType, effect.EffectInformation.BuildByLevelEffect, effect.Exit));
            }
            if (items.Count > 0)
            {
                list.InsertItem(new BlockItem(clickNo, items));
            }
            return list;
        }

        private void Handle(SlideRange sldRange)
        {
            if (PowerPointCurrentPresentationInfo.CurrentSlide != null)
            {
                blockItemList = InitializeBlockItemList();
                listView.ItemsSource = null;
                listView.ItemsSource = BlockItems;
            }
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            blockDragManager = new ListViewDragDropManager<BlockItem>(listView);
            listView.Drop += ListView_Drop;
        }

        private void ListView_Drop(object sender, DragEventArgs e)
        {
            for (int i = 0; i < this.listView.Items.Count; ++i)
            {
                ListViewItem item = GetListViewItem(listView, i);
                Label label = GetChildOfType<Label>(item);
                if (label != null)
                {
                    label.Content = i.ToString();
                }
            }
            CustomAnimationItem data = e.Data.GetData(typeof(CustomAnimationItem)) as CustomAnimationItem;
            if (data == null)
            {
                return;
            }
            for (int i = 0; i < this.listView.Items.Count; ++i)
            {
                ListViewItem item = GetListViewItem(listView, i);
                ListView view = GetChildOfType<ListView>(item);
                if (IsMouseDirectOver(view) && draggedListView != view)
                {
                    ObservableCollection<AnimationItem> list = draggedListView.ItemsSource as ObservableCollection<AnimationItem>;
                    list.Remove(data);
                    if (list.Count() == 0)
                    {
                        (listView.ItemsSource as ObservableCollection<BlockItem>).RemoveAt(draggedListViewIndex);
                    }
                }               
            }         
        }

        private void ListView_PreviewMouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            for (int i = 0; i < this.listView.Items.Count; ++i)
            {
                ListViewItem item = GetListViewItem(listView, i);
                ListView view = GetChildOfType<ListView>(item);
                if (IsMouseDirectOver(view))
                {
                    draggedListView = view;
                    draggedListViewIndex = i;
                }              
            }
        }

        ListViewItem GetListViewItem(ListView listview, int index)
        {
            if (listView.ItemContainerGenerator.Status != GeneratorStatus.ContainersGenerated)
            {
                return null;
            }
            return listView.ItemContainerGenerator.ContainerFromIndex(index) as ListViewItem;
        }
        private ListView GetChildOfType<ListView>(DependencyObject depObj) where ListView : DependencyObject
        {
            if (depObj == null)
            {
                return null;
            }
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
            {
                var child = VisualTreeHelper.GetChild(depObj, i);

                var result = (child as ListView) ?? GetChildOfType<ListView>(child);
                if (result != null)
                {
                    return result;
                }
            }
            return null;
        }

        private bool IsMouseDirectOver(Visual target)
        {
            Rect bounds = VisualTreeHelper.GetDescendantBounds(target);
            System.Windows.Point mousePos = MouseUtilities.GetMousePosition(target);
            return bounds.Contains(mousePos);
        }
    }
}
