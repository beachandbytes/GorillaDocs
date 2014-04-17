using GorillaDocs.ViewModels;

namespace GorillaDocs.Views
{
    public partial class AddTableView : OfficeDialog
    {
        readonly AddTableViewModel viewModel = null;

        public AddTableView() { InitializeComponent(); }

        public AddTableView(AddTableViewModel viewModel)
            : this()
        {
            this.viewModel = viewModel;
            this.DataContext = this.viewModel;
        }
    }
}
