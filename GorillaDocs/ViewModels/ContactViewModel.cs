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

        public ContactViewModel(Contact contact, List<string> deliveryItems, Outlook outlook, string examplePhoneNumber = "")
        {
            Contact = contact;
            Outlook = outlook;
            ExamplePhoneNumber = examplePhoneNumber;

            AddressBookCommand = new GenericCommand(AddressBookPressed);
            ClearCommand = new GenericCommand(ClearPressed, CanClear);

            DeliveryItems = deliveryItems;
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

        Contact contact;
        public Contact Contact
        {
            //TODO: Merge the ContactViewModel and SenderViewModel code
            get { return contact; }
            set
            {
                if (contact != null && value != null && contact.FullName == value.FullName)
                    return;

                if (value == null)
                    contact = new Contact();
                else if (contact == null)
                    contact = value;
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
        public bool CanClear()
        {
            return !Contact.IsNullOrEmpty();
        }
        [LoudRibbonExceptionHandler]
        public void ClearPressed() { Contact.Clear(); }
    }
}
