using GorillaDocs.ViewModels;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Linq;
using System.Xml.Serialization;

namespace GorillaDocs.Models
{
    public class ContactComparer : IEqualityComparer<Contact>
    {
        public bool Equals(Contact x, Contact y) { return x.FullName == y.FullName; }
        public int GetHashCode(Contact obj) { return obj.FullName.GetHashCode(); }
    }

    [Serializable]
    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    [XmlRoot("Favourites")]
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
                Replace(contact);
            return contact;
        }

        public new void Add(Contact contact)
        {
            if (this.Any(x => x.FullName == contact.FullName))
                throw new InvalidOperationException(string.Format("Can not add duplicate '{0}'", contact.FullName));
            base.Add(contact);
            NotifyPropertyChanged("Favourites");
        }

        public new void Remove(Contact contact)
        {
            this.RemoveAll(x => x.FullName == contact.FullName);
            NotifyPropertyChanged("Favourites");
        }

        public new bool Contains(Contact contact)
        {
            return this.Any(x => contact != null && x.FullName == contact.FullName);
        }

        public void Replace(Contact contact)
        {
            var temp = this.First(x => x.FullName == contact.FullName);
            var i = this.IndexOf(temp);
            this.RemoveAt(i);
            this.Insert(i, contact);
        }

        public new event PropertyChangedEventHandler PropertyChanged;
        protected void NotifyPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
