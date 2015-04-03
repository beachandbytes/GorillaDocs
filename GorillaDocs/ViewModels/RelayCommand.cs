using GorillaDocs.libs.PostSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;

namespace GorillaDocs.ViewModels
{
    [Log]
    public class RelayCommand : ICommand
    {
        readonly Action executeMethod;
        readonly Func<bool> canExecuteMethod;

        public RelayCommand(Action executeMethod, Func<bool> canExecuteMethod = null)
        {
            this.executeMethod = executeMethod;
            this.canExecuteMethod = canExecuteMethod;
        }

        public bool CanExecute(object parameter)
        {
            if (canExecuteMethod == null)
                return true;
            return canExecuteMethod();
        }

        public void Execute(object parameter) { executeMethod(); }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
    }
}
