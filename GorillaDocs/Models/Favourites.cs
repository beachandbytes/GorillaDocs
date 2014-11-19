using GorillaDocs.ViewModels;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;

namespace GorillaDocs.Models
{
    public class ContactComparer : IEqualityComparer<Contact>
    {
        public bool Equals(Contact x, Contact y) { return x.Fullname == y.Fullname; }
        public int GetHashCode(Contact obj) { return obj.Fullname.GetHashCode(); }
    }

    public class Favourites : ObservableCollection<Contact>
    {
        //TODO: Should I use OnCollectionChanged or OnPropertyChanged instead of my own NotifyPropertyChanged?

        public Favourites(ObservableCollection<Contact> contacts)
        {
            if (contacts != null)
                foreach (Contact contact in contacts.Distinct(new ContactComparer()))
                    Add(contact);
        }

        public Contact FirstOrPassedIn(Contact contact)
        {
            if (Contains(contact))
                return this.First(x => x.Fullname == contact.Fullname);
            else
                return contact;
        }

        public new void Add(Contact contact)
        {
            if (this.Any(x => x.Fullname == contact.Fullname))
                throw new InvalidOperationException(string.Format("Can not add duplicate '{0}'", contact.Fullname));
            base.Add(contact);
            NotifyPropertyChanged("Favourites");
        }

        public new void Remove(Contact contact)
        {
            this.RemoveAll(x => x.Fullname == contact.Fullname);
            NotifyPropertyChanged("Favourites");
        }

        public new bool Contains(Contact contact)
        {
            return this.Any(x => contact != null && x.Fullname == contact.Fullname);
        }

        public new event PropertyChangedEventHandler PropertyChanged;
        protected void NotifyPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
