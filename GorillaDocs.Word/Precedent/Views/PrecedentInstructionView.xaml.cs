using GorillaDocs.Views;
using System;
using System.Windows.Controls;
using System.Windows.Threading;
using GorillaDocs.Word.Precedent.ViewModels;

namespace GorillaDocs.Word.Precedent.Views
{
    public partial class PrecedentInstructionView : OfficeDialog
    {
        public PrecedentInstructionViewModel viewModel = null;

        public PrecedentInstructionView() { InitializeComponent(); }

        public PrecedentInstructionView(PrecedentInstructionViewModel viewModel)
            : this()
        {
            this.viewModel = viewModel;
            DataContext = this.viewModel;
        }

        public void Grid_Error(object sender, ValidationErrorEventArgs e)
        {
            if (e.Action == ValidationErrorEventAction.Added) PrecedentInstructionViewModel.Errors += 1;
            if (e.Action == ValidationErrorEventAction.Removed) PrecedentInstructionViewModel.Errors -= 1;
        }

        public void Grid_SetInitialFocus(object sender, System.Windows.DependencyPropertyChangedEventArgs e)
        {
            if ((bool)e.NewValue == true)
                Dispatcher.BeginInvoke(
                        DispatcherPriority.ContextIdle,
                        new Action(delegate() { Command.Focus(); }));
        }
    }
}
