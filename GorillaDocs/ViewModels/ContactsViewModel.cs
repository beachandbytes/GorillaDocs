using GorillaDocs.libs.PostSharp;
using GorillaDocs.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace GorillaDocs.ViewModels
{
    /// <summary>
    /// WARNING.. This class is complicated because there are 2 collections (Contacts and Favourites) that intermingle when data bound. 
    /// </summary>
    [Log]
    public class ContactsViewModel : Notify
    {
        readonly int MaxContacts;

        public ObservableCollection<Contact> Contacts { get; set; }
        readonly Outlook Outlook;

        public ContactsViewModel(ObservableCollection<Contact> contacts, List<string> deliveryItems, Outlook outlook, int maxContacts, Favourites favourites)
        {
            Favourites = new Favourites(new ObservableCollection<Contact>());
            //Favourites = favourites;
            Favourites.PropertyChanged += Favourites_PropertyChanged;
            MaxContacts = maxContacts;
            Contacts = contacts;
            Outlook = outlook;

            AddressBookCommand = new GenericCommand(AddressBookPressed);
            ClearCommand = new GenericCommand(ClearPressed, CanClear);
            NextCommand = new GenericCommand(NextPressed, CanPressNext);
            PrevCommand = new GenericCommand(PrevPressed, CanPressPrev);
            AddFavouriteCommand = new GenericCommand(AddFavouritePressed, CanAddFavourite);
            RemoveFavouriteCommand = new GenericCommand(RemoveFavouritePressed);

            DeliveryItems = deliveryItems;
            Contact = favourites.FirstOrPassedIn(Contacts.FirstOrCreateIfEmpty());
        }

        public List<string> DeliveryItems { get; private set; }

        public Favourites Favourites { get; private set; }
        Contact contact;
        public Contact Contact
        {
            get { return contact; }
            set
            {
                if (contact != null && value != null && contact.FullName == value.FullName)
                    return;

                if (value == null)
                    contact = new Contact();
                else if (contact != null && Favourites.Contains(value) && !Contacts.Contains(value))
                    contact.Copy(value);
                else
                    contact = value;
                NotifyPropertyChanged("Contact");
                NotifyPropertyChanged("IsEnabled");
                NotifyPropertyChanged("AddFavouriteVisibility");
                NotifyPropertyChanged("RemoveFavouriteVisibility");
                NotifyPropertyChanged("Count");
            }
        }

        public string Count
        {
            get
            {
                var list = Contacts.ToList();
                list.RemoveAll(x => x.IsEmpty());
                int current = list.FindIndex(x => x.FullName == Contact.FullName) + 1;
                int count = list.Count();
                if (current == 0)
                    return string.Format("{0} of {0}", count + 1);
                else
                    return string.Format("{0} of {1}", current, count);
            }
        }

        public ICommand AddressBookCommand { get; set; }
        public void AddressBookPressed()
        {
            try
            {
                Contact ol = Outlook.GetContact();
                if (ol != null)
                {
                    ClearPressed();
                    Contact.Copy(ol);
                }
                //Contact = Contacts.ReplaceAndReturn(Contact, ol);
            }
            catch (Exception ex)
            {
                Message.ShowError(ex);
            }
        }

        public ICommand ClearCommand { get; set; }
        public bool CanClear()
        {
            return !Contact.IsNullOrEmpty();
        }
        public void ClearPressed()
        {
            try
            {
                int i = IndexOf(Contact);
                //if (i == -1 || i == Contacts.Count)
                if (Contact == Contacts.Last())
                {
                    if (Favourites.Contains(Contact))
                        Contact = new Contact();
                    Contact.Clear();
                }
                else
                {
                    Contacts.RemoveAll(x => x.FullName == this.Contact.FullName);
                    Contact = Contacts[i];
                }
            }
            catch (Exception ex)
            {
                Message.ShowError(ex);
            }
        }

        public ICommand PrevCommand { get; set; }
        public bool CanPressPrev() { return IndexOf(Contact) > 0; }
        public void PrevPressed()
        {
            try
            {
                int i = IndexOf(Contact);
                Contact = Contacts[i - 1];
            }
            catch (Exception ex)
            {
                Message.ShowError(ex);
            }
        }

        public ICommand NextCommand { get; set; }
        public bool CanPressNext() { return !Contact.IsNullOrEmpty() && IndexOf(Contact) != MaxContacts; }
        public void NextPressed()
        {
            try
            {
                if (IndexOf(Contact) + 1 == Contacts.Count)
                {
                    Contact = new Contact();
                    Contacts.Add(Contact);
                }
                else
                {
                    int i = IndexOf(Contact);
                    Contact = Contacts[i + 1];
                }
            }
            catch (Exception ex)
            {
                Message.ShowError(ex);
            }
        }

        public ICommand AddFavouriteCommand { get; set; }
        public bool CanAddFavourite()
        {
            return false;
            //return this.Contact != null && !string.IsNullOrEmpty(this.Contact.Fullname);
        }
        public void AddFavouritePressed()
        {
            try
            {
                Favourites.Add(Contact);
            }
            catch (Exception ex)
            {
                Message.ShowError(ex);
            }
        }

        public ICommand RemoveFavouriteCommand { get; set; }
        public void RemoveFavouritePressed()
        {
            try
            {
                var temp = Contact;
                Favourites.Remove(Contact); // binding causes Contact to become null
                Contact = temp;
            }
            catch (Exception ex)
            {
                Message.ShowError(ex);
            }
        }

        void Favourites_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            try
            {
                Properties.Settings.Default.Favourites = new ContactCollection(Favourites);
                Properties.Settings.Default.Save();
                NotifyPropertyChanged("IsEnabled");
                NotifyPropertyChanged("AddFavouriteVisibility");
                NotifyPropertyChanged("RemoveFavouriteVisibility");
            }
            catch (Exception ex)
            {
                Message.ShowError(ex);
            }
        }

        public bool IsEnabled { get { return !Favourites.Contains(Contact); } }
        public Visibility AddFavouriteVisibility { get { return IsEnabled ? Visibility.Visible : Visibility.Collapsed; } }
        public Visibility RemoveFavouriteVisibility { get { return IsEnabled ? Visibility.Collapsed : Visibility.Visible; } }
        int IndexOf(Contact contact)
        {
            return Contacts.ToList().FindIndex(x => x.FullName == contact.FullName);
        }
    }
}
