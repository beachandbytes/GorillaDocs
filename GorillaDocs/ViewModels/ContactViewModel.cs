using GorillaDocs.libs.PostSharp;
using GorillaDocs.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;

namespace GorillaDocs.ViewModels
{
    public class ContactViewModel : Notify
    {
        string _ExamplePhoneNumber;
        readonly Outlook Outlook;

        public ContactViewModel(Contact contact, Outlook outlook, string examplePhoneNumber = "", List<string> deliveryItems = null)
        {
            this.contact = contact;
            Outlook = outlook;
            ExamplePhoneNumber = examplePhoneNumber;

            AddressBookCommand = new RelayCommand(AddressBookPressed);
            ClearCommand = new RelayCommand(ClearPressed, CanClear);

            DeliveryItems = deliveryItems ?? new List<string>();
        }

        public string ExamplePhoneNumber
        {
            get { return _ExamplePhoneNumber; }
            set
            {
                _ExamplePhoneNumber = value;
                NotifyPropertyChanged("ExamplePhoneNumber");
            }
        }

        public List<string> DeliveryItems { get; private set; }

        readonly Contact contact;
        public Contact Contact
        {
            get { return contact; }
            set
            {
                if (value == null)
                    contact.Clear();
                else
                    contact.Copy(value);

                if (string.IsNullOrEmpty(contact.Delivery))
                    contact.Delivery = DeliveryItems.FirstOrDefault();

                NotifyPropertyChanged("Contact");
            }
        }

        public ICommand AddressBookCommand { get; set; }
        [LoudRibbonExceptionHandler]
        public void AddressBookPressed()
        {
            Contact ol = Outlook.GetContact();
            if (ol != null)
            {
                ClearPressed();
                Contact = ol;
            }
        }

        public ICommand ClearCommand { get; set; }
        public bool CanClear() { return !Contact.IsNullOrEmpty(); }
        [LoudRibbonExceptionHandler]
        public void ClearPressed() { Contact.Clear(); }
    }
}
