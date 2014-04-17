﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel;

namespace GorillaDocs.ViewModels
{
    public class BaseViewModel : INotifyPropertyChanged
    {
        public BaseViewModel() { }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void NotifyPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

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