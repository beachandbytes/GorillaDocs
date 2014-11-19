using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;

namespace GorillaDocs.ViewModels
{
    public class ListBoxContact : Contact
    {
        readonly ListBoxContacts _Parent;

        public ListBoxContact(ListBoxContacts parent)
        {
            _Parent = parent;
            EditCommand = new GenericCommand(EditPressed);
            RemoveCommand = new GenericCommand(RemovePressed);
        }

        bool _IsEditMode;
        public bool IsEditMode
        {
            get { return _IsEditMode; }
            set
            {
                _IsEditMode = value;
                NotifyPropertyChanged("IsEditMode");
            }
        }

        public ICommand EditCommand { get; set; }
        public void EditPressed()
        {
            IsEditMode = !IsEditMode;
        }

        public ICommand RemoveCommand { get; set; }
        public void RemovePressed()
        {
            _Parent.Remove(this);
        }
    }
}
