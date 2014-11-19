using GorillaDocs.libs.PostSharp;
using GorillaDocs.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace GorillaDocs.ViewModels
{
    [Log]
    public class SenderViewModel : Notify
    {
        public SenderViewModel(Contact contact, Outlook outlook, Favourites favourites)
        {
            if (contact == null)
                throw new ArgumentNullException("Contact");
            if (outlook == null)
                throw new ArgumentNullException("Outlook");

            Favourites = favourites;
            Favourites.PropertyChanged += Favourites_PropertyChanged;
            Contact = favourites.FirstOrPassedIn(contact);
            Outlook = outlook;

            AddressBookCommand = new GenericCommand(AddressBookPressed);
            ClearCommand = new GenericCommand(ClearPressed, CanClear);
            AddFavouriteCommand = new GenericCommand(AddFavouritePressed, CanAddFavourite);
            RemoveFavouriteCommand = new GenericCommand(RemoveFavouritePressed);
        }

        public Contact Contact
        {
            get { return _Contact; }
            set
            {
                _Contact = value == null ? new Contact() : value;
                NotifyPropertyChanged("Contact");
                NotifyPropertyChanged("IsEnabled");
                NotifyPropertyChanged("AddFavouriteVisibility");
                NotifyPropertyChanged("RemoveFavouriteVisibility");
            }
        }
        Contact _Contact;
        readonly Outlook Outlook;

        public Favourites Favourites { get; private set; }

        public ICommand AddressBookCommand { get; set; }
        public void AddressBookPressed()
        {
            try
            {
                ClearPressed();
                Contact ol = this.Outlook.GetContact();
                if (ol != null)
                    Contact.Copy(ol);
            }
            catch (Exception ex)
            {
                Message.ShowError(ex);
            }
        }

        public ICommand ClearCommand { get; set; }
        public bool CanClear()
        {
            return !this.Contact.IsNullOrEmpty();
        }
        public void ClearPressed()
        {
            try
            {
                if (this.Favourites.Contains(Contact))
                    this.Contact = new Contact();
                this.Contact.Clear();
            }
            catch (Exception ex)
            {
                Message.ShowError(ex);
            }
        }

        public ICommand AddFavouriteCommand { get; set; }
        public bool CanAddFavourite()
        {
            return this.Contact != null && !string.IsNullOrEmpty(this.Contact.Fullname);
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
    }
}
