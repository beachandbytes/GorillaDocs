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
    public class ContactWithFavouritesViewModel : Notify
    {
        public ContactWithFavouritesViewModel(Contact contact, Outlook outlook, Favourites favourites)
        {
            if (contact == null)
                throw new ArgumentNullException("Contact");
            if (outlook == null)
                throw new ArgumentNullException("Outlook");

            Outlook = outlook;

            Contact = contact;
            Contact.PropertyChanged += Contact_PropertyChanged;

            Favourites = favourites;
            Favourites.PropertyChanged += Favourites_PropertyChanged;
            //Contact = favourites.FirstOrDefault(x => x.FullName == contact.FullName);
            //ParentContact = contact;
            //if (Contact.IsEmpty() && !ParentContact.IsEmpty())
            //    Contact.Copy(ParentContact);
            //Contact.PropertyChanged += Contact_PropertyChanged;

            AddressBookCommand = new RelayCommand(AddressBookPressed);
            ClearCommand = new RelayCommand(ClearPressed, CanClear);
            AddFavouriteCommand = new RelayCommand(AddFavouritePressed, CanAddFavourite);
            RemoveFavouriteCommand = new RelayCommand(RemoveFavouritePressed);
        }

        bool disablePropertyChangedEvent = false;
        void Contact_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (!disablePropertyChangedEvent && e.PropertyName == "FullName" && Favourites != null && Favourites.Any(x => x.FullName == Contact.FullName))
            {
                try
                {
                    disablePropertyChangedEvent = true;
                    var user = Favourites.First(x => x.FullName == Contact.FullName);
                    Contact.Copy(user);
                }
                finally
                {
                    disablePropertyChangedEvent = false;
                }
            }
        }

        //readonly Contact ParentContact; // Need to maintain a separate contact so that Favourites list doesn't reset it
        readonly Contact Contact;
        //public Contact Contact
        //{
        //    get { return _Contact; }
        //    set
        //    {
        //        if (value == null)
        //            ClearPressed();
        //        else
        //        {
        //            _Contact = value;
        //            if (ParentContact != null)
        //                ParentContact.Copy(value);
        //        }

        //        NotifyPropertyChanged("Contact");
        //        NotifyPropertyChanged("IsEnabled");
        //        NotifyPropertyChanged("AddFavouriteVisibility");
        //        NotifyPropertyChanged("RemoveFavouriteVisibility");
        //    }
        //}
        readonly Outlook Outlook;

        public Favourites Favourites { get; private set; }

        public ICommand AddressBookCommand { get; set; }
        public void AddressBookPressed()
        {
            try
            {
                Contact ol = this.Outlook.GetContact();
                if (ol != null)
                {
                    ClearPressed();
                    Contact.Copy(ol);
                }
            }
            catch (Exception ex)
            {
                Message.ShowError(ex);
            }
        }

        public ICommand ClearCommand { get; set; }
        public bool CanClear() { return !Contact.IsNullOrEmpty(); }
        public void ClearPressed()
        {
            try
            {
                //if (Favourites.Contains(Contact))
                //    Contact = new Contact();
                //else
                //    Contact.Clear();
                //if (ParentContact != null)
                //ParentContact.Clear();
                Contact.Clear();
            }
            catch (Exception ex)
            {
                Message.ShowError(ex);
            }
        }

        public ICommand AddFavouriteCommand { get; set; }
        public bool CanAddFavourite() { return Contact != null && !string.IsNullOrEmpty(Contact.FullName); }
        public void AddFavouritePressed()
        {
            try
            {
                var temp = new Contact();
                temp.Copy(Contact);
                Favourites.Add(temp);
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
                Favourites.RemoveAll(x => x.FullName == Contact.FullName);
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

        public bool IsEnabled { get { return !Favourites.Any(x => x.FullName == Contact.FullName); } }
        public Visibility AddFavouriteVisibility { get { return IsEnabled ? Visibility.Visible : Visibility.Collapsed; } }
        public Visibility RemoveFavouriteVisibility { get { return IsEnabled ? Visibility.Collapsed : Visibility.Visible; } }
    }
}
