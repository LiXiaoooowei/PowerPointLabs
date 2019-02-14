﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
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

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;
using PowerPointLabs.ELearningLab.ELearningWorkspace.ModelFactory;
using PowerPointLabs.ELearningLab.Extensions;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.ELearningLab.Utility;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ELearningLab.Views
{
    /// <summary>
    /// Interaction logic for ELearningLabMainPanel.xaml
    /// </summary>
    public partial class ELearningLabMainPanel : UserControl
    {
        public ObservableCollection<ClickItem> Items { get; set; }
        private PowerPointSlide slide;
        public int FirstClickNumber
        {
            get
            {
                return slide.IsFirstAnimationTriggeredByClick() ? 1 : 0;
            }
        }
        private ELearningLabMainPanel()
        {
            slide = this.GetCurrentSlide();
            InitializeComponent();
            Items = LoadItems();
            listView.ItemsSource = Items;
            syncImage.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
               Properties.Resources.Refresh.GetHbitmap(),
               IntPtr.Zero,
               Int32Rect.Empty,
               BitmapSizeOptions.FromEmptyOptions());
        }

        private static ELearningLabMainPanel instance;
        public static ELearningLabMainPanel GetInstance()
        {
            if (instance == null)
            {
                instance = new ELearningLabMainPanel();
                return instance;
            }
            return instance;
        }

        public void HandleELearningPaneVisibilityChanged()
        {
            slide = this.GetCurrentSlide();
            Items = LoadItems();
            listView.ItemsSource = Items;
        }

        public void RefreshVoiceLabelOnAudioSettingChanged()
        {
            if (Visibility == Visibility.Visible)
            {
                ObservableCollection<ClickItem> clickItems = listView.ItemsSource as ObservableCollection<ClickItem>;
                foreach (ClickItem item in clickItems)
                {
                    if (item is SelfExplanationClickItem)
                    {
                        SelfExplanationClickItem selfExplanationClickItem = item as SelfExplanationClickItem;
                        if (StringUtility.ExtractDefaultLabelFromVoiceLabel(selfExplanationClickItem.VoiceLabel)
                            .Equals(ELearningLabText.DefaultAudioIdentifier))
                        {
                            selfExplanationClickItem.VoiceLabel = string.Format(ELearningLabText.AudioDefaultLabelFormat,
                                AudioSettingService.selectedVoice.ToString());
                        }
                    }
                }
            }
        }

        public void SyncClickItems()
        {
            CheckAzureAccountValidity();
            SyncCustomAnimationToTaskpane();
            RefreshListViewItemsSource();
            RemoveLabAnimationsFromAnimationPane();
            SyncLabItemToAnimationPane();
        }

        public bool IsInSync()
        {
            try
            {
                ObservableCollection<ClickItem> items_loaded = LoadItems();
                ObservableCollection<ClickItem> items_original = Items;
                return items_loaded.SequenceEqual(items_original);
            }
            catch
            {
                Logger.Log("exception in sync");
                return true;
            }
        }

        private ObservableCollection<ClickItem> LoadItems()
        {
            SelfExplanationTagService.Clear();
            int clickNo = FirstClickNumber;
            ObservableCollection<ClickItem> clickBlocks = new ObservableCollection<ClickItem>();
            ClickItem customClickBlock, selfExplanationClickBlock;
            do
            {
                customClickBlock =
                    new CustomItemFactory(slide.GetCustomEffectsForClick(clickNo), slide).GetBlock();
                selfExplanationClickBlock =
                    new SelfExplanationItemFactory(slide.GetPPTLEffectsForClick(clickNo), slide).GetBlock();

                if (customClickBlock != null)
                {
                    customClickBlock.ClickNo = clickNo;
                    clickBlocks.Add(customClickBlock);
                }
                if (selfExplanationClickBlock != null)
                {
                    selfExplanationClickBlock.ClickNo = clickNo;
                    if (customClickBlock == null && selfExplanationClickBlock is SelfExplanationClickItem) // is independent block
                    {
                        (selfExplanationClickBlock as SelfExplanationClickItem).TriggerIndex = (int)TriggerType.OnClick;
                    }
                    clickBlocks.Add(selfExplanationClickBlock);
                }
                clickNo++;
            }
            while (customClickBlock != null || selfExplanationClickBlock != null);

            return clickBlocks;
        }

        #region Custom Event Handlers for SelfExplanationBlockView

        private void HandleUpButtonClickedEvent(object sender, RoutedEventArgs e)
        {
            SelfExplanationClickItem labItem = ((Button)e.OriginalSource).CommandParameter as SelfExplanationClickItem;
            int index = Items.IndexOf(labItem);
            if (index > 0)
            {
                Items.Move(index, index - 1);
            }
            RefreshListViewItemsSource();
        }
        private void HandleDownButtonClickedEvent(object sender, RoutedEventArgs e)
        {
            SelfExplanationClickItem labItem = ((Button)e.OriginalSource).CommandParameter as SelfExplanationClickItem;
            int index = Items.IndexOf(labItem);
            if (index < Items.Count() - 1 && index >= 0)
            {
                Items.Move(index, index + 1);
            }
            RefreshListViewItemsSource();
        }
        private void HandleDeleteButtonClickedEvent(object sender, RoutedEventArgs e)
        {
            SelfExplanationClickItem labItem = ((Button)e.OriginalSource).CommandParameter as SelfExplanationClickItem;
            Items.Remove(labItem);
            RefreshListViewItemsSource();
        }
        private void HandleTriggerTypeComboBoxSelectionChangedEvent(object sender, RoutedEventArgs e)
        {
            RefreshListViewItemsSource();
        }

        #endregion

        #region XMAL-Binded Event Handler

        private void SyncButton_Click(object sender, RoutedEventArgs e)
        {
            SyncClickItems();     
        }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            string text = textBox.Text.Trim();
            if (string.IsNullOrEmpty(text))
            {
                MessageBox.Show("Text must not be empty!");
                return;
            }
            SelfExplanationClickItem selfExplanationClickItem = new SelfExplanationClickItem(captionText: text);
            selfExplanationClickItem.tagNo = SelfExplanationTagService.GenerateUniqueTag();
            (listView.ItemsSource as ObservableCollection<ClickItem>).Add(selfExplanationClickItem);
            textBox.Text = string.Empty;
        }

        #endregion

        #region Helper Methods

        private void SyncCustomAnimationToTaskpane()
        {
            Queue<CustomClickItem> customClickItems = LoadCustomClickItems();
            ReplaceCustomItemsInItemsSource(customClickItems);
            UpdateClickNoInItemsSource();
        } 

        private void SyncLabItemToAnimationPane()
        {

            ELearningService.SyncLabItemToAnimationPane(slide, 
                Items.Where(
                    x => x is SelfExplanationClickItem).Cast<SelfExplanationClickItem>().ToList());
        }

        private void RemoveLabAnimationsFromAnimationPane()
        {
            slide.RemoveAnimationsForShapeWithPrefix(ELearningLabText.Identifier);
        }

        /// <summary>
        /// Load custom animations from animation pane separated by click
        /// </summary>
        /// <returns>Queue of CustomClickItem</returns>
        private Queue<CustomClickItem> LoadCustomClickItems()
        {
            int clickNo = FirstClickNumber;
            Queue<CustomClickItem> customClickItems = new Queue<CustomClickItem>();
            ClickItem customClickBlock;
            do
            {
                customClickBlock =
                    new CustomItemFactory(slide.GetCustomEffectsForClick(clickNo), slide).GetBlock();
                
                if (customClickBlock != null)
                {
                    customClickBlock.ClickNo = clickNo;
                    customClickItems.Enqueue((CustomClickItem)customClickBlock);
                }
    
                clickNo++;
            }
            while (slide.TimeLine.MainSequence.FindFirstAnimationForClick(clickNo) != null);

            return customClickItems;
        }

        /// <summary>
        /// Refresh list view, force ClickNo label to update by `BlockToIndexConverter`
        /// </summary>
        private void RefreshListViewItemsSource()
        {
            ICollectionView view = CollectionViewSource.GetDefaultView(listView.ItemsSource);
            view.Refresh();
        }

        /// <summary>
        /// Replace all CustomClickItem in ItemsSource with `customClickItems`
        /// If there are more CustomClickItem in ItemsSource, those are deleted.
        /// Additional customClickItem left in customClickItems will be appended to the back of list.
        /// </summary>
        /// <param name="customClickItems"></param>
        /// <returns></returns>
        private ObservableCollection<ClickItem> ReplaceCustomItemsInItemsSource(Queue<CustomClickItem> customClickItems)
        {
            for (int i = 0; i < Items.Count(); i++)
            {
                ClickItem clickItem = Items.ElementAt(i);
                if (clickItem is CustomClickItem)
                {
                    if (customClickItems.Count() > 0)
                    {
                        CustomClickItem customClickItem = customClickItems.Dequeue();
                        Items.RemoveAt(i);
                        Items.Insert(i, customClickItem);
                    }
                    else
                    {
                        Items.RemoveAt(i);
                        i--;
                    }
                }
            }
            while (customClickItems.Count() > 0)
            {
                CustomClickItem customClickItem = customClickItems.Dequeue();
                Items.Add(customClickItem);
            }
            return Items;
        }

        /// <summary>
        /// Update ClickNo property in ClickItem when old CustomClickItem is replaced with new ones.
        /// Note that we cannot reply on `BlockToIndexConverter` to update ClickNo, 
        /// because ListViewItem which invokes `BlockToIndexConverter`, is loaded lazily.
        /// </summary>
        /// <param name="clickItems"></param>
        /// <returns></returns>
        private ObservableCollection<ClickItem> UpdateClickNoInItemsSource()
        {
            int clickNo = FirstClickNumber;
            for (int i = 0; i < Items.Count(); i++)
            {
                ClickItem clickItem = Items.ElementAt(i);
                if (i == 0)
                {
                    clickItem.ClickNo = clickNo;
                }
                else if (clickItem is SelfExplanationClickItem && 
                    Items.ElementAt(i - 1) is CustomClickItem && 
                    (clickItem as SelfExplanationClickItem).TriggerIndex != (int)TriggerType.OnClick)
                {
                    clickItem.ClickNo = Items.ElementAt(i - 1).ClickNo;
                }
                else
                {
                    clickItem.ClickNo = Items.ElementAt(i - 1).ClickNo + 1;
                }
            }
            return Items;
        }

        private void CheckAzureAccountValidity()
        {
            AzureAccountStorageService.LoadUserAccount();
            if (!AzureRuntimeService.IsAzureAccountPresent() || !AzureRuntimeService.IsValidUserAccount())
            {
                MessageBox.Show("Azure Account Authentication Failed. \n Azure Voices Cannot Be Generated.");
            }
        }

        #endregion
    }
}
