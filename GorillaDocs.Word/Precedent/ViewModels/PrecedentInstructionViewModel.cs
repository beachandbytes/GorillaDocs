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
            this.Details = Details;
            this.Namespaces = Namespaces;
            if (control.Type == Wd.WdContentControlType.wdContentControlComboBox)
                IsComboBox = true;
            if (control.Type == Wd.WdContentControlType.wdContentControlDate && control.DateStorageFormat != Wd.WdContentControlDateStorageFormat.wdContentControlDateStorageDate)
                control.DateStorageFormat = Wd.WdContentControlDateStorageFormat.wdContentControlDateStorageDate;
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
