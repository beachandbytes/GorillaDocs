using System;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel;
using GorillaDocs.libs.PostSharp;

namespace GorillaDocs.ViewModels
{
    [Log]
    public class BaseViewModel : Notify
    {
        public BaseViewModel() { }

        bool? dialogResult;
        public bool? DialogResult
        {
            get { return this.dialogResult; }
            set
            {
                this.dialogResult = value;
                NotifyPropertyChanged("DialogResult");
            }
        }
    }
}
