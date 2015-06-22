using GorillaDocs.ViewModels;
using System;
using System.Runtime.InteropServices;
using System.Windows.Input;

namespace GorillaDocs.Views
{
    public partial class SelectTemplateView : OfficeDialog
    {
        public readonly SelectTemplateViewModel viewModel = null;

        public SelectTemplateView() { InitializeComponent(); }
        public SelectTemplateView(SelectTemplateViewModel viewModel)
            : this()
        {
            this.viewModel = viewModel;
            DataContext = this.viewModel;
        }

        void lstAllTemplates_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (viewModel.CanPressOK())
                viewModel.OKPressed();
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsWindow(IntPtr hWnd);
    }
}
