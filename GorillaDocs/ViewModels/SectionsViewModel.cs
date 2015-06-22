namespace GorillaDocs.ViewModels
{
    public class SectionsViewModel : BaseViewModel
    {
        public SectionsViewModel() { this.OKCommand = new RelayCommand(OKPressed, CanPressOK); }

        public bool FrontCover { get; set; }
        public bool ExecutiveSummary { get; set; }
        public bool TableOfContents { get; set; }
        public bool NormalBlankPage { get; set; }
        public bool Appendix { get; set; }
        public bool BackCover { get; set; }
        public bool AgreementBody { get; set; }

        public RelayCommand OKCommand { get; set; }
        public bool CanPressOK() { return FrontCover || ExecutiveSummary || TableOfContents || NormalBlankPage || Appendix || BackCover || AgreementBody; }
        public void OKPressed() { this.DialogResult = true; }
    }
}
