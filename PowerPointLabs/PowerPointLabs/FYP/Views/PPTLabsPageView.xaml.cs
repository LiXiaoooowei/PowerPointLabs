using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
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

using PowerPointLabs.FYP.Data;
using PowerPointLabs.FYP.Service;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.FYP.Views
{
#pragma warning disable 0618
    /// <summary>
    /// Interaction logic for PPTLabsPageView.xaml
    /// </summary>
    public partial class PPTLabsPageView : UserControl
    {
        
        public ObservableCollection<AnimationItem> Items { get; private set; }
        private ListViewDragDropManager<AnimationItem> itemDragManager;

        public PPTLabsPageView()
        {
            if (LicenseManager.UsageMode == LicenseUsageMode.Runtime)
            {
                InitializeComponent();
                Items = InitializeBlockItemList();
                Globals.ThisAddIn.Application.SlideSelectionChanged += Handle;
                labListView.ItemsSource = Items;
            }
        }
        internal void HandleSyncButtonClick()
        {
            throw new NotImplementedException();
        }

        private ObservableCollection<AnimationItem> InitializeBlockItemList()
        {
            LabAnimationItemIdentifierManager.EmptyTagsCollection();
            IEnumerable<Effect> effects = PowerPointCurrentPresentationInfo.CurrentSlide.TimeLine.MainSequence.Cast<Effect>();
            ObservableCollection<AnimationItem> items = new ObservableCollection<AnimationItem>();
            Dictionary<int, LabAnimationItem> labItems = new Dictionary<int, LabAnimationItem>();
            for (int i = 0; i < effects.Count(); i++)
            {
                Effect effect = effects.ElementAt(i);
                
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
                        }
                        if (isCallout)
                        {
                            labItem.IsCallout = true;
                        }
                        if (isVoice)
                        {
                            labItem.IsVoice = true;
                        }
                    }
                    else
                    {
                        labItem = new LabAnimationItem(effect.Shape.TextFrame.TextRange.Text, tagNo, isCaption, isVoice, isCallout);
                        labItems.Add(tagNo, labItem);
                        items.Add(labItem);
                    }
                }
                else
                {
                    items.Add(new CustomAnimationItem(effect.Shape,
                               effect.EffectType, effect.EffectInformation.BuildByLevelEffect, effect.Exit));
                }
               
            }
            return items;
        }

        private void Handle(SlideRange sldRange)
        {
            if (PowerPointCurrentPresentationInfo.CurrentSlide != null)
            {
                Items = InitializeBlockItemList();
                labListView.ItemsSource = null;
                labListView.ItemsSource = Items;
            }
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            itemDragManager = new ListViewDragDropManager<AnimationItem>(labListView);
        }
    }
}
