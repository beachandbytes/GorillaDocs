using GorillaDocs.ViewModels;

namespace GorillaDocs.Views
{
    public partial class ChangeOfficeView : OfficeDialog
    {
        readonly ChangeOfficeViewModel viewModel = null;

        public ChangeOfficeView() { InitializeComponent(); }

        public ChangeOfficeView(ChangeOfficeViewModel viewModel)
            : this()
        {
            this.viewModel = viewModel;
            this.DataContext = this.viewModel;
        }
    }
}
