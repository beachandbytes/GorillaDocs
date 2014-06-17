using GorillaDocs.libs.PostSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;

namespace GorillaDocs.ViewModels
{
    [Log]
    public class AddTableViewModel : BaseViewModel
    {
        public AddTableViewModel()
        {
            NumberOfColumns = 5;
            NumberOfRows = 2;
            this.OKCommand = new GenericCommand(OKPressed);
        }

        public int NumberOfColumns { get; set; }
        public int NumberOfRows { get; set; }
        public string TableHeading { get; set; }
        public string TableSource { get; set; }

        public ICommand OKCommand { get; set; }
        public void OKPressed() { this.DialogResult = true; }
    }
}
