using System.ComponentModel;
using PowerPointLabs.FYP.Data;
using PowerPointLabs.NarrationsLab.Data;

namespace PowerPointLabs.NarrationsLab.Views
{
    /// <summary>
    /// Interaction logic for NarrationsLabSettingsDialogBox.xaml
    /// </summary>
    public partial class NarrationsLabSettingsDialogBox : INotifyPropertyChanged
    {


        public event PropertyChangedEventHandler PropertyChanged = (sender, e) => { };
        public NarrationsLabSettingsPage CurrentPage
        {
            get
            {
                return _currentPage;
            }
            set
            {
                _currentPage = value;
                PropertyChanged(this, new PropertyChangedEventArgs("CurrentPage"));
            }
        }
        public LabAnimationItem labAnimationItem;

        private static NarrationsLabSettingsDialogBox instance;
        
        private NarrationsLabSettingsPage _currentPage { get; set; } = NarrationsLabSettingsPage.MainSettingsPage;
        public void SetCurrentPage(NarrationsLabSettingsPage page)
        {
            CurrentPage = page;
        }

        public static NarrationsLabSettingsDialogBox GetInstance(
            NarrationsLabSettingsPage page = NarrationsLabSettingsPage.MainSettingsPage,
            LabAnimationItem item = null)
        {
            if (instance == null)
            {
                instance = new NarrationsLabSettingsDialogBox();
                instance._currentPage = page;
                instance.labAnimationItem = item;
            }
            return instance;
        }
        public void Destroy()
        {
            HumanVoiceLoginPage.GetInstance().Destroy();
            NarrationsLabMainSettingsPage.GetInstance().Destroy();
            instance = null;
        }
        private NarrationsLabSettingsDialogBox()
        {
            InitializeComponent();
            this.DataContext = this;
        }
    }
}
