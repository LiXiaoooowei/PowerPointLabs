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
        public ObservableCollection<AnimationItem> Items { get; set; }
        private ListView draggedListView = null;
        private int draggedListViewIndex = -1;
        //   private ListViewDragDropManager<BlockItem> blockDragManager;


        public BlockPageView()
        {
            if (LicenseManager.UsageMode == LicenseUsageMode.Runtime)
            {
                InitializeComponent();
                Items = InitializeItemList();
                Globals.ThisAddIn.Application.SlideSelectionChanged += HandleSlideSelectionChange;
                listView.ItemsSource = Items;
            }
        }

        public void HandleSyncButtonClick()
        {
            PowerPointSlide slide = PowerPointCurrentPresentationInfo.CurrentSlide;
            IEnumerable<Effect> effects = slide.TimeLine.MainSequence.Cast<Effect>();
            IEnumerable<Shape> shapes = slide.Shapes.Cast<Shape>();
            foreach (Shape shape in slide.Shapes)
            {
                if (shape.Name.Contains(FYPText.Identifier))
                {
                    slide.RemoveAnimationsForShape(shape);
                }
            }
            AddAppearanceLabAnimation();
            AddDisappearanceLabAnimation();
        }

        public void AddLabAnimationItem(LabAnimationItem item)
        {
            ObservableCollection<AnimationItem> items = listView.ItemsSource as ObservableCollection<AnimationItem>;
            AnimationItem lastItem = items.ElementAt(items.Count() - 1);
            if (lastItem is CustomAnimationItems)
            {
                item.IsTailEnabled = true;
            }
            else
            {
                item.IsTailEnabled = false;
            }
            items.Add(item);
        }

        private void AddAppearanceLabAnimation()
        {
            ObservableCollection<AnimationItem> animationItems = listView.ItemsSource as ObservableCollection<AnimationItem>;
            PowerPointSlide slide = PowerPointCurrentPresentationInfo.CurrentSlide;
            IEnumerable<Effect> effects = slide.TimeLine.MainSequence.Cast<Effect>();

            for (int i = 0; i < listView.Items.Count; ++i)
            {
                ListViewItem listViewItem = GetListViewItem(listView, i);
                Label label = GetChildOfType<Label>(listViewItem);
                if (label != null && animationItems.ElementAt(i) is LabAnimationItem)
                {
                    int clickNo = Convert.ToInt32(label.Content.ToString());

                    LabAnimationItem item = animationItems.ElementAt(i) as LabAnimationItem;
                    if (item.IsTail && item.IsTailEnabled) // tail lab item
                    {
                        SyncLabAnimationItemToSlide(item as LabAnimationItem, slide, clickNo);
                    }
                    else // independent lab item
                    {
                        SyncLabAnimationItemToSlide(item as LabAnimationItem, slide, clickNo - 1, isSeperateClick: true);
                    }

                }
            }
        }

        private void AddDisappearanceLabAnimation()
        {
            ObservableCollection<AnimationItem> animationItems = listView.ItemsSource as ObservableCollection<AnimationItem>;
            PowerPointSlide slide = PowerPointCurrentPresentationInfo.CurrentSlide;
            IEnumerable<Effect> effects = slide.TimeLine.MainSequence.Cast<Effect>();

            for (int i = 0; i < listView.Items.Count; ++i)
            {
                ListViewItem listViewItem = GetListViewItem(listView, i);
                Label label = GetChildOfType<Label>(listViewItem);
                if (label != null && animationItems.ElementAt(i) is LabAnimationItem)
                {
                    int clickNo = Convert.ToInt32(label.Content.ToString());
                    LabAnimationItem item = animationItems.ElementAt(i) as LabAnimationItem;
                    SyncLabAnimationItemToSlide(item as LabAnimationItem, slide, clickNo, isSeperateClick: false, syncAppearance: false);
                }
            }
        }

        private void SyncLabAnimationItemToSlide(LabAnimationItem item, PowerPointSlide slide, int clickNo, 
            bool isSeperateClick = false, bool syncAppearance = true)
        {
            item.Execute(slide, clickNo, isSeperateClick, syncAppearance);
        }

        private ObservableCollection<AnimationItem> InitializeItemList()
        {
            LabAnimationItemIdentifierManager.EmptyTagsCollection();
            IEnumerable<Effect> effects = PowerPointCurrentPresentationInfo.CurrentSlide.TimeLine.MainSequence.Cast<Effect>();
            ObservableCollection<AnimationItem> list = new ObservableCollection<AnimationItem>();
            ObservableCollection<CustomAnimationItem> customItems = new ObservableCollection<CustomAnimationItem>();
            Dictionary<int, LabAnimationItem> labItems = new Dictionary<int, LabAnimationItem>();
            LabAnimationItem labItem = null;
            PowerPointSlide slide = PowerPointCurrentPresentationInfo.CurrentSlide;
            int clickNo = slide.IsFirstAnimationTriggeredByClick() ? 1 : 0;
            bool prevBlkContainsNoLabItem = false;
            for (int i = 0; i < effects.Count(); i++)
            {
                Effect effect = effects.ElementAt(i);
                if (slide.TimeLine.MainSequence.FindFirstAnimationForClick(clickNo) == effect)
                {
                    bool isTail = false;
                    if (customItems.Count() != 0)
                    {
                        list.Add(new CustomAnimationItems(customItems, clickNo - 1));
                        customItems = new ObservableCollection<CustomAnimationItem>();
                        isTail = true;
                    }
                    if (labItem != null)
                    {
                        labItem.IsTail = isTail;
                        if (prevBlkContainsNoLabItem || isTail)
                        {
                            labItem.IsTailEnabled = true;
                        }
                        list.Add(labItem);
                        labItem = null;
                    }
                    else
                    {
                        prevBlkContainsNoLabItem = true;
                    }
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
                    if (labItem != null)
                    {
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
                            labItem = new LabAnimationItem(clickNo - 1, text, tagNo, note, isCaption, isVoice, isCallout, voiceLabel: voiceName);
                        }
                        catch (Exception e)
                        {
                            Logger.Log(e.Message);
                        }
                    }
                }
                else
                {
                    customItems.Add(new CustomAnimationItem(effect.Shape,
                               effect.EffectType, effect.EffectInformation.BuildByLevelEffect, effect.Exit));
                }
            }
            if (customItems.Count() != 0)
            {
                list.Add(new CustomAnimationItems(customItems, clickNo));
                customItems = new ObservableCollection<CustomAnimationItem>();
            }
            if (labItem != null)
            {
                list.Add(labItem);
                labItem = null;
            }

            return list;
        }

        private void HandleSlideSelectionChange(SlideRange sldRange)
        {
            if (PowerPointCurrentPresentationInfo.CurrentSlide != null)
            {
                Items = InitializeItemList();
                listView.ItemsSource = null;
                listView.ItemsSource = Items;
            }
        }

        private void HandleUpBtnClickedEvent(object sender, RoutedEventArgs e)
        {
            LabAnimationItem labItem = ((Button)e.OriginalSource).CommandParameter as LabAnimationItem;
            ObservableCollection<AnimationItem> blockItems = listView.ItemsSource as ObservableCollection<AnimationItem>;
            int index = blockItems.IndexOf(labItem);
            if (index > 0)
            {
                blockItems.Move(index, index - 1);
            }
            ICollectionView view = CollectionViewSource.GetDefaultView(listView.ItemsSource);
            view.Refresh();
        }

        private void HandleDownBtnClickedEvent(object sender, RoutedEventArgs e)
        {
            LabAnimationItem labItem = ((Button)e.OriginalSource).CommandParameter as LabAnimationItem;
            ObservableCollection<AnimationItem> blockItems = listView.ItemsSource as ObservableCollection<AnimationItem>;
            int index = blockItems.IndexOf(labItem);
            if (index < blockItems.Count() - 1 && index >= 0)
            {
                blockItems.Move(index, index + 1);
            }
            ICollectionView view = CollectionViewSource.GetDefaultView(listView.ItemsSource);
            view.Refresh();
        }

        private void HandleTailCheckedEvent(object sender, RoutedEventArgs e)
        {
            LabAnimationItem labItem = ((CheckBox)e.OriginalSource).CommandParameter as LabAnimationItem;
            ObservableCollection<AnimationItem> blockItems = listView.ItemsSource as ObservableCollection<AnimationItem>;
            int index = blockItems.IndexOf(labItem);
            for (int i = index; i < blockItems.Count() && i >= 0; i++)
            {
                blockItems.ElementAt(i).ClickNo -= 1;
                Label label = GetChildOfType<Label>(GetListViewItem(listView, i));
                label.Content = blockItems.ElementAt(i).ClickNo.ToString();
            }
        }

        private void HandleTailUncheckedEvent(object sender, RoutedEventArgs e)
        {
            LabAnimationItem labItem = ((CheckBox)e.OriginalSource).CommandParameter as LabAnimationItem;
            ObservableCollection<AnimationItem> blockItems = listView.ItemsSource as ObservableCollection<AnimationItem>;
            int index = blockItems.IndexOf(labItem);
            for (int i = index; i < blockItems.Count() && i >= 0; i++)
            {
                blockItems.ElementAt(i).ClickNo += 1;
                Label label = GetChildOfType<Label>(GetListViewItem(listView, i));
                label.Content = blockItems.ElementAt(i).ClickNo.ToString();
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
