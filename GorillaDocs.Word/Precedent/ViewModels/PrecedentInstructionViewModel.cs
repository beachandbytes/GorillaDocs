using GorillaDocs.libs.PostSharp;
using GorillaDocs.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.Precedent.ViewModels
{
    [Log]
    public class PrecedentInstructionViewModel : BaseViewModel
    {
        public PrecedentInstructionViewModel() { }
        public PrecedentInstructionViewModel(PrecedentInstruction Details, Wd.ContentControl control, IList<string> Namespaces)
        {
            //TODO: Enforce 'Date' content control data type. Ensure not 'Date and Time'
            this.Details = Details;
            this.Namespaces = Namespaces;
            if (control.Type == Wd.WdContentControlType.wdContentControlComboBox)
                IsComboBox = true;
            OKCommand = new RelayCommand(OKPressed, CanPressOK);
        }

        public PrecedentInstruction Details { get; set; }
        public IList<string> Namespaces { get; set; }
        public bool IsComboBox { get; set; }

        public static int Errors { get; set; }

        public RelayCommand OKCommand { get; set; }
        public void OKPressed() { DialogResult = true; }
        public bool CanPressOK() { return Errors == 0; }
    }
}
