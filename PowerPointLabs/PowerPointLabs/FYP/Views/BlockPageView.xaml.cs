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
using PowerPointLabs.TextCollection;

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

        public void HandleSyncButtonClick()
        {
            PowerPointSlide slide = PowerPointCurrentPresentationInfo.CurrentSlide;
            IEnumerable<Effect> effects = slide.TimeLine.MainSequence.Cast<Effect>();
            IEnumerable<Shape> shapes = slide.Shapes.Cast<Shape>();
            // For old version, uncomment the line below
            // slide.RemoveAnimationsForShapes(shapes.ToList());
            // For new version
            foreach (Shape shape in slide.Shapes)
            {
                if (shape.Name.Contains(FYPText.Identifier))
                {
                    slide.RemoveAnimationsForShape(shape);
                }
            }
            //end for new version
            AddAppearanceLabAnimation();
            AddDisappearanceLabAnimation();
        }

        public void AddLabAnimationItem(LabAnimationItem item)
        {
            (listView.ItemsSource as ObservableCollection<BlockItem>)
                .Add(new BlockItem(-1, new ObservableCollection<AnimationItem>() { item }));
        }

        private void AddAppearanceLabAnimation()
        {
            ObservableCollection<BlockItem> animationItems = listView.ItemsSource as ObservableCollection<BlockItem>;
            PowerPointSlide slide = PowerPointCurrentPresentationInfo.CurrentSlide;
            IEnumerable<Effect> effects = slide.TimeLine.MainSequence.Cast<Effect>();

            for (int i = 0; i < listView.Items.Count; ++i)
            {
                ListViewItem listViewItem = GetListViewItem(listView, i);
                Label label = GetChildOfType<Label>(listViewItem);
                if (label != null)
                {
                    int clickNo = Convert.ToInt32(label.Content.ToString());
                    BlockItem blockItem = animationItems.ElementAt(i);
                    bool containsCustomAnimationInBlock = ContainsCustomAnimationInBlock(blockItem);
                    if (i == 0 && effects.Count() > 0)
                    {
                        effects.ElementAt(0).Timing.TriggerType = containsCustomAnimationInBlock && clickNo == 0? MsoAnimTriggerType.msoAnimTriggerWithPrevious:
                            MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                    }

                    for (int j = 0; j < blockItem.Items.Count; j++)
                    {
                        AnimationItem item = blockItem.Items.ElementAt(j) as AnimationItem;
                        if (item.GetType() == typeof(LabAnimationItem) && containsCustomAnimationInBlock)
                        {
                            SyncLabAnimationItemToSlide(item as LabAnimationItem, slide, clickNo, j);
                        }
                        else if (item.GetType() == typeof(LabAnimationItem) && !containsCustomAnimationInBlock)
                        {
                            if ((i == 0 && clickNo == 0) || j != 0)
                            {
                                SyncLabAnimationItemToSlide(item as LabAnimationItem, slide, clickNo, j, false);
                            }
                            else if (j == 0)
                            {
                                SyncLabAnimationItemToSlide(item as LabAnimationItem, slide, clickNo - 1, j, true);
                            }
                            else
                            {
                                SyncLabAnimationItemToSlide(item as LabAnimationItem, slide, clickNo - 1, j, false);
                            }
                        }
                    }
                }
            }
        }

        private void AddDisappearanceLabAnimation()
        {
            ObservableCollection<BlockItem> animationItems = listView.ItemsSource as ObservableCollection<BlockItem>;
            PowerPointSlide slide = PowerPointCurrentPresentationInfo.CurrentSlide;
            IEnumerable<Effect> effects = slide.TimeLine.MainSequence.Cast<Effect>();

            for (int i = 0; i < listView.Items.Count; ++i)
            {
                ListViewItem listViewItem = GetListViewItem(listView, i);
                Label label = GetChildOfType<Label>(listViewItem);
                if (label != null)
                {
                    int clickNo = Convert.ToInt32(label.Content.ToString());
                    BlockItem blockItem = animationItems.ElementAt(i);
                    for (int j = 0; j < blockItem.Items.Count; j++)
                    {
                        AnimationItem item = blockItem.Items.ElementAt(j) as AnimationItem;
                        if (item.GetType() == typeof(LabAnimationItem))
                        {
                            SyncLabAnimationItemToSlide(item as LabAnimationItem, slide, clickNo, j, false, false);
                        }                        
                    }
                }
            }
        }

        private bool ContainsCustomAnimationInBlock(BlockItem blockItem)
        {
            for (int j = 0; j < blockItem.Items.Count; j++)
            {
                AnimationItem item = blockItem.Items.ElementAt(j) as AnimationItem;
                if (item.GetType() == typeof(CustomAnimationItem))
                {
                    return true;
                }
            }
            return false;
        }

        private void SyncCustomAnimationItemToSlide(CustomAnimationItem item, PowerPointSlide slide, int clickNo, int j)
        {
            Effect effect = slide.SetShapeAsClickTriggered(item.GetShape(), clickNo, item.GetEffectType());
            if (item.GetExit() == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                effect.Exit = Microsoft.Office.Core.MsoTriState.msoTrue;               
            }
        }
         
        private void SyncLabAnimationItemToSlide(LabAnimationItem item, PowerPointSlide slide, int clickNo, int seqNo, 
            bool isSeperateClick = false, bool syncAppearance = true)
        {
            item.Execute(slide, clickNo, seqNo, isSeperateClick, syncAppearance);
        }

        private BlockItemList InitializeBlockItemList()
        {
            LabAnimationItemIdentifierManager.EmptyTagsCollection();
            IEnumerable<Effect> effects = PowerPointCurrentPresentationInfo.CurrentSlide.TimeLine.MainSequence.Cast<Effect>();
            BlockItemList list = new BlockItemList();
            ObservableCollection<AnimationItem> items = new ObservableCollection<AnimationItem>();
            Dictionary<int, LabAnimationItem> labItems = new Dictionary<int, LabAnimationItem>();
            int clickNo = PowerPointCurrentPresentationInfo.CurrentSlide.IsFirstAnimationTriggeredByClick() ? 1 : 0;
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
                if (LabAnimationItemIdentifierManager.GetTagNo(effect.Shape.Name) != -1)
                {
                    if (effect.Exit == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        continue;
                    }
                    int tagNo = LabAnimationItemIdentifierManager.GetTagNo(effect.Shape.Name);
                    string functionMatch = LabAnimationItemIdentifierManager.GetTagFunction(effect.Shape.Name);
                    bool isCaption = functionMatch == FYPText.CaptionIdentifier;
                    bool isCallout = functionMatch == FYPText.CalloutIdentifier;
                    bool isVoice = functionMatch == FYPText.AudioIdentifier;
                    LabAnimationItem labItem;
                    if (labItems.ContainsKey(tagNo))
                    {
                        labItem = labItems[tagNo];
                        if (isCaption)
                        {
                            labItem.IsCaption = true;
                            labItem.Note = effect.Shape.TextFrame.TextRange.Text.Trim();
                        }
                        if (isCallout)
                        {
                            labItem.IsCallout = true;
                            labItem.Text = effect.Shape.TextFrame.TextRange.Text.Trim();
                        }
                        if (isVoice)
                        {
                            string voiceName = LabAnimationItemIdentifierManager.GetVoiceName(effect.Shape.Name); 
                            labItem.IsVoice = isVoice;
                            Shape shape = PowerPointCurrentPresentationInfo.CurrentSlide
                                .GetShapeWithName(FYPText.Identifier + FYPText.Underscore + tagNo.ToString() + FYPText.Underscore + FYPText.CaptionIdentifier)[0];
                            labItem.Note = shape.TextFrame.TextRange.Text.Trim();
                            labItem.VoiceLabel = voiceName;
                        }
                    }
                    else
                    {
                        try
                        {
                            string text = isCallout ? effect.Shape.TextFrame.TextRange.Text.Trim() : "";
                            Shape shape = PowerPointCurrentPresentationInfo.CurrentSlide
                                .GetShapeWithName(FYPText.Identifier + FYPText.Underscore + tagNo.ToString() + FYPText.Underscore + FYPText.CaptionIdentifier)[0];
                            string note = shape.TextFrame.TextRange.Text.Trim();
                            string voiceName = LabAnimationItemIdentifierManager.GetVoiceName(effect.Shape.Name); 
                            labItem = new LabAnimationItem(text, tagNo, note, isCaption, isVoice, isCallout, voiceLabel: voiceName);
                            labItems.Add(tagNo, labItem);
                            items.Add(labItem);
                        }
                        catch (Exception e)
                        {
                            Logger.Log(e.Message);
                        }
                    }
                }
                else
                {
                    items.Add(new CustomAnimationItem(effect.Shape,
                               effect.EffectType, effect.EffectInformation.BuildByLevelEffect, effect.Exit));
                }
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

        private void HandleUpBtnClickedEvent(object sender, RoutedEventArgs e)
        {
            LabAnimationItem labItem = ((Button)e.OriginalSource).CommandParameter as LabAnimationItem;
            ObservableCollection<BlockItem> blockItems = listView.ItemsSource as ObservableCollection<BlockItem>;
            for (int i = 0; i < listView.Items.Count; i++)
            {
                ListViewItem item = GetListViewItem(listView, i);
                ListView view = GetChildOfType<ListView>(item);
                if (IsMouseDirectOver(view))
                {
                    ObservableCollection<AnimationItem> list = view.ItemsSource as ObservableCollection<AnimationItem>;
                    list.Remove(labItem);
                    if (i > 0 && list.Count == 0)
                    {
                        ListViewItem prevItem = GetListViewItem(listView, i - 1);
                        ListView prevView = GetChildOfType<ListView>(prevItem);
                        ObservableCollection<AnimationItem> prevList = prevView.ItemsSource as ObservableCollection<AnimationItem>;
                        prevList.Add(labItem);
                    }
                    else
                    {
                        blockItems.Insert(i, new BlockItem(0, new ObservableCollection<AnimationItem>() { labItem}));
                    }
                    if (list.Count() == 0)
                    {
                        (listView.ItemsSource as ObservableCollection<BlockItem>).RemoveAt(i);
                    }
                    break;
                }
            }
            for (int i = 0; i < this.listView.Items.Count; ++i)
            {
                ListViewItem item = GetListViewItem(listView, i);
                Label label = GetChildOfType<Label>(item);
                if (label != null)
                {
                    label.Content = PowerPointCurrentPresentationInfo.CurrentSlide.IsFirstAnimationTriggeredByClick() ? (i + 1).ToString() : i.ToString();
                }
            }
        }

        private void HandleDownBtnClickedEvent(object sender, RoutedEventArgs e)
        {
            LabAnimationItem labItem = ((Button)e.OriginalSource).CommandParameter as LabAnimationItem;
            ObservableCollection<BlockItem> blockItems = listView.ItemsSource as ObservableCollection<BlockItem>;
            for (int i = 0; i < listView.Items.Count; i++)
            {
                ListViewItem item = GetListViewItem(listView, i);
                ListView view = GetChildOfType<ListView>(item);
                if (IsMouseDirectOver(view))
                {
                    ObservableCollection<AnimationItem> list = view.ItemsSource as ObservableCollection<AnimationItem>;
                    list.Remove(labItem);
                    if (i < listView.Items.Count - 1 && list.Count == 0)
                    {
                        ListViewItem nextItem = GetListViewItem(listView, i + 1);
                        ListView nextView = GetChildOfType<ListView>(nextItem);
                        ObservableCollection<AnimationItem> nextList = nextView.ItemsSource as ObservableCollection<AnimationItem>;
                        nextList.Insert(0, labItem);
                    }
                    else
                    {
                        blockItems.Insert(i + 1, new BlockItem(0, new ObservableCollection<AnimationItem>() { labItem }));
                    }
                    if (list.Count() == 0)
                    {
                        (listView.ItemsSource as ObservableCollection<BlockItem>).RemoveAt(i);
                    }
                    break;
                }
            }
            for (int i = 0; i < this.listView.Items.Count; ++i)
            {
                ListViewItem item = GetListViewItem(listView, i);
                Label label = GetChildOfType<Label>(item);
                if (label != null)
                {
                    label.Content = PowerPointCurrentPresentationInfo.CurrentSlide.IsFirstAnimationTriggeredByClick() ? (i + 1).ToString() : i.ToString();
                }
            }
        }

        private void ListView_Drop(object sender, DragEventArgs e)
        {
            // for non-dragged listview items
            for (int i = 0; i < this.listView.Items.Count; ++i)
            {
                ListViewItem item = GetListViewItem(listView, i);
                Label label = GetChildOfType<Label>(item);
                if (label != null)
                {
                    label.Content = PowerPointCurrentPresentationInfo.CurrentSlide.IsFirstAnimationTriggeredByClick() ? (i + 1).ToString() : i.ToString();
                }
            }
            AnimationItem data = e.Data.GetDataPresent(typeof(CustomAnimationItem)) ?
                e.Data.GetData(typeof(CustomAnimationItem)) as AnimationItem :
                e.Data.GetData(typeof(LabAnimationItem)) as AnimationItem;
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
