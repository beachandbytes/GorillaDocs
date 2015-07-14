namespace GorillaDocs.ViewModels
{
    public class SectionsViewModel : BaseViewModel
    {
        public SectionsViewModel() { this.OKCommand = new RelayCommand(OKPressed, CanPressOK); }

        bool _FrontCover;
        bool _LightFrontCover;
        bool _DarkFrontCover;
        bool _ExecutiveSummary;
        bool _TableOfContents;
        bool _NormalBlankPage;
        bool _Appendix;
        bool _BackCover;
        bool _LightBackCover;
        bool _DarkBackCover;
        bool _AgreementBody;

        public bool FrontCover
        {
            get { return _FrontCover; }
            set
            {
                _FrontCover = value;
                NotifyPropertyChanged("FrontCover");
            }
        }
        public bool LightFrontCover
        {
            get { return _LightFrontCover; }
            set
            {
                _LightFrontCover = value;
                NotifyPropertyChanged("LightFrontCover");
            }
        }
        public bool DarkFrontCover
        {
            get { return _DarkFrontCover; }
            set
            {
                _DarkFrontCover = value;
                NotifyPropertyChanged("DarkFrontCover");
            }
        }
        public bool ExecutiveSummary
        {
            get { return _ExecutiveSummary; }
            set
            {
                _ExecutiveSummary = value;
                NotifyPropertyChanged("ExecutiveSummary");
            }
        }
        public bool TableOfContents
        {
            get { return _TableOfContents; }
            set
            {
                _TableOfContents = value;
                NotifyPropertyChanged("TableOfContents");
            }
        }
        public bool NormalBlankPage
        {
            get { return _NormalBlankPage; }
            set
            {
                _NormalBlankPage = value;
                NotifyPropertyChanged("NormalBlankPage");
            }
        }
        public bool Appendix
        {
            get { return _Appendix; }
            set
            {
                _Appendix = value;
                NotifyPropertyChanged("Appendix");
            }
        }
        public bool BackCover
        {
            get { return _BackCover; }
            set
            {
                _BackCover = value;
                NotifyPropertyChanged("BackCover");
            }
        }
        public bool LightBackCover
        {
            get { return _LightBackCover; }
            set
            {
                _LightBackCover = value;
                NotifyPropertyChanged("LightBackCover");
            }
        }
        public bool DarkBackCover
        {
            get { return _DarkBackCover; }
            set
            {
                _DarkBackCover = value;
                NotifyPropertyChanged("DarkBackCover");
            }
        }
        public bool AgreementBody
        {
            get { return _AgreementBody; }
            set
            {
                _AgreementBody = value;
                NotifyPropertyChanged("AgreementBody");
            }
        }

        public RelayCommand OKCommand { get; set; }
        public bool CanPressOK() { return FrontCover || LightFrontCover || DarkFrontCover || ExecutiveSummary || TableOfContents || NormalBlankPage || Appendix || BackCover || LightBackCover || DarkBackCover || AgreementBody; }
        public void OKPressed() { this.DialogResult = true; }
    }
}
