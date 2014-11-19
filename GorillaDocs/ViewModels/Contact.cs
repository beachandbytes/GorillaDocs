using GorillaDocs.libs.PostSharp;
using GorillaDocs.Properties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Xml.Serialization;

namespace GorillaDocs.ViewModels
{
    [Log]
    public class Contact : Notify, IEquatable<Contact>
    {
        //TODO: Tidy INotifyPropertyChanged code
        //http://www.codeproject.com/Articles/38865/INotifyPropertyChanged-auto-wiring-or-how-to-get-rid-of-redundant-code.aspx
        //http://stackoverflow.com/questions/1315621/implementing-inotifypropertychanged-does-a-better-way-exist
        //https://github.com/Fody/PropertyChanged
        string _Initials;
        string _Fullname;
        string _FirstName;
        string _LastName;
        string _Title;
        string _CompanyName;
        string _PhoneNumber;
        string _FaxNumber;
        string _EmailAddress;
        string _Address;
        string _Country;
        string _Delivery;

        public string Initials
        {
            get { return _Initials; }
            set
            {
                _Initials = value;
                NotifyPropertyChanged("Initials");
            }
        }
        public string Fullname
        {
            get { return _Fullname; }
            set
            {
                _Fullname = value;
                NotifyPropertyChanged("Fullname");
            }
        }
        public string FirstName
        {
            get { return _FirstName; }
            set
            {
                _FirstName = value;
                NotifyPropertyChanged("FirstName");
            }
        }
        public string LastName
        {
            get { return _LastName; }
            set
            {
                _LastName = value;
                NotifyPropertyChanged("LastName");
            }
        }
        public string Title
        {
            get { return _Title; }
            set
            {
                _Title = value;
                NotifyPropertyChanged("Title");
            }
        }
        public string CompanyName
        {
            get { return _CompanyName; }
            set
            {
                _CompanyName = value;
                NotifyPropertyChanged("CompanyName");
            }
        }
        public string PhoneNumber
        {
            get { return _PhoneNumber; }
            set
            {
                _PhoneNumber = value;
                NotifyPropertyChanged("PhoneNumber");
            }
        }
        [XmlElement(IsNullable = true)]
        public string FaxNumber
        {
            get { return _FaxNumber; }
            set
            {
                _FaxNumber = value;
                NotifyPropertyChanged("FaxNumber");
            }
        }
        [XmlElement(IsNullable = true)]
        public string EmailAddress
        {
            get { return _EmailAddress; }
            set
            {
                _EmailAddress = value;
                NotifyPropertyChanged("EmailAddress");
            }
        }
        [XmlElement(IsNullable = true)]
        public string Address
        {
            get { return _Address; }
            set
            {
                _Address = value;
                NotifyPropertyChanged("Address");
            }
        }
        public string Country
        {
            get { return _Country; }
            set
            {
                _Country = value;
                NotifyPropertyChanged("Country");
            }
        }
        public string Delivery
        {
            get { return _Delivery; }
            set
            {
                _Delivery = value;
                NotifyPropertyChanged("Delivery");
                NotifyPropertyChanged("EmailVisibility");
                NotifyPropertyChanged("FaxVisibility");
                NotifyPropertyChanged("AddressVisibility");
            }
        }

        public bool IsDeliveryByEmail { get { return Delivery != null && Delivery.Contains(Resources.Delivery_Email); } }
        public bool IsDeliveryByFax { get { return Delivery != null && Delivery.Contains(Resources.Delivery_Facsimile); } }
        public bool IsDeliveryByOther { get { return Delivery != null && Delivery.Contains(Resources.Delivery_Other); } }

        public Visibility EmailVisibility
        {
            get
            {
                if (Delivery != null && Delivery.Contains(Resources.Delivery_Email))
                    return Visibility.Visible;
                return Visibility.Collapsed;
            }
        }
        public Visibility FaxVisibility
        {
            get
            {
                if (Delivery != null && Delivery.Contains(Resources.Delivery_Facsimile))
                    return Visibility.Visible;
                return Visibility.Collapsed;
            }
        }
        public Visibility AddressVisibility
        {
            get
            {
                if (Delivery != null && Delivery.Contains(Resources.Delivery_Other))
                    return Visibility.Visible;
                return Visibility.Collapsed;
            }
        }

        public bool IsEmpty()
        {
            if (!string.IsNullOrEmpty(Initials))
                return false;
            if (!string.IsNullOrEmpty(Fullname))
                return false;
            if (!string.IsNullOrEmpty(FirstName))
                return false;
            if (!string.IsNullOrEmpty(LastName))
                return false;
            if (!string.IsNullOrEmpty(Title))
                return false;
            if (!string.IsNullOrEmpty(CompanyName))
                return false;
            if (!string.IsNullOrEmpty(PhoneNumber))
                return false;
            if (!string.IsNullOrEmpty(FaxNumber))
                return false;
            if (!string.IsNullOrEmpty(EmailAddress))
                return false;
            if (!string.IsNullOrEmpty(Address))
                return false;
            if (!string.IsNullOrEmpty(Country))
                return false;
            if (!string.IsNullOrEmpty(Delivery))
                return false;
            return true;
        }

        public void Clear()
        {
            Initials = string.Empty;
            Fullname = string.Empty;
            FirstName = string.Empty;
            LastName = string.Empty;
            Title = string.Empty;
            CompanyName = string.Empty;
            PhoneNumber = string.Empty;
            FaxNumber = string.Empty;
            EmailAddress = string.Empty;
            Address = string.Empty;
            Country = string.Empty;
            Delivery = string.Empty;
        }
        public void Copy(Contact from)
        {
            Initials = from.Initials;
            Fullname = from.Fullname;
            FirstName = from.FirstName;
            LastName = from.LastName;
            Title = from.Title;
            CompanyName = from.CompanyName;
            PhoneNumber = from.PhoneNumber;
            FaxNumber = from.FaxNumber;
            EmailAddress = from.EmailAddress;
            Address = from.Address;
            Country = from.Country;
            Delivery = from.Delivery;
        }
        public bool Equals(Contact other)
        {
            if (other == null)
                return false;
            if (other == this)
                return true;
            return false;
        }

        public override bool Equals(object obj) { return Equals(obj as Contact); }

        public override string ToString() { return Fullname; }
    }
}
