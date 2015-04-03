using GorillaDocs.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;

namespace GorillaDocs.ViewModels
{
    public class ListBoxContacts : Notify
    {
        public ListBoxContacts()
        {
            Title = "Recipients";
            Contacts = new ObservableCollection<Contact>();
            AddCommand = new RelayCommand(AddPressed);

            Contacts.Add(new ListBoxContact(this) { FirstName = "Matthew", LastName = "Fitzmaurice", FullName = "Matthew Fitzmaurice" });
            Contacts.Add(new ListBoxContact(this) { FirstName = "Marcia", LastName = "Fitzmaurice", FullName = "Marcia Fitzmaurice" });
        }

        public string Title { get; set; }

        public ObservableCollection<Contact> Contacts { get; set; }

        public ICommand AddCommand { get; set; }
        public void AddPressed()
        {
            Contacts.Add(new ListBoxContact(this) { IsEditMode = true });
        }

        public void Remove(ListBoxContact contact)
        {
            Contacts.Remove(contact);
        }
    }
}
