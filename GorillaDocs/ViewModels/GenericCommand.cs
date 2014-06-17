using GorillaDocs.libs.PostSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;

namespace GorillaDocs.ViewModels
{
    [Log]
    public class GenericCommand : ICommand
    {
        public delegate void ExecuteMethod();
        public delegate bool CanExecuteMethod();
        readonly ExecuteMethod executeMethod;
        readonly CanExecuteMethod canExecuteMethod;

        public GenericCommand(ExecuteMethod executeMethod, CanExecuteMethod canExecuteMethod = null)
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

        public void Execute(object parameter)
        {
            executeMethod();
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
    }
}
