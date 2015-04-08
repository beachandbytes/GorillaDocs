using DataAnnotationsExtensions;
using GorillaDocs.libs.PostSharp;
using GorillaDocs.Properties;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Xml.Serialization;

namespace GorillaDocs.Models
{
    [Log]
    public class Contact : EntityBase, IEquatable<Contact>
    {
        string _Title;
        string _Initials;
        string _FullName;
        string _FirstName;
        string _LastName;
        string _Salutation;
        string _Qualifications;
        string _Position;
        string _CompanyName;
        string _Department;
        string _PhoneNumber;
        string _FaxNumber;
        string _EmailAddress;
        string _Address;
        string _StreetAddress1;
        string _StreetAddress2;
        string _StreetCity;
        string _StreetState;
        string _StreetPostalCode;
        string _StreetCountry;
        string _PostalAddress1;
        string _PostalAddress2;
        string _PostalCity;
        string _PostalState;
        string _PostalPostalCode;
        string _PostalCountry;
        string _Country;
        string _Delivery;
        string _SignatureLine1;
        string _SignatureLine2;
        List<string> _States;
        List<string> _Titles;

        bool CanUpdateSalutation()
        {
            return string.IsNullOrEmpty(Salutation) || Salutation == Title || Salutation == LastName || Salutation == String.Format("{0} {1}", _Title, LastName).Trim();
        }
        public string Title
        {
            get { return _Title; }
            set
            {
                if (CanUpdateSalutation())
                    Salutation = String.Format("{0} {1}", value, LastName).Trim();
                _Title = value;
                NotifyPropertyChanged("Title");
                NotifyPropertyChanged("Salutation");
            }
        }
        public string Initials
        {
            get { return _Initials; }
            set
            {
                _Initials = value;
                NotifyPropertyChanged("Initials");
            }
        }
        public string FullName
        {
            get { return _FullName; }
            set
            {
                _FullName = value;
                NotifyPropertyChanged("FullName");
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
                if (CanUpdateSalutation())
                    Salutation = String.Format("{0} {1}", Title, value).Trim();
                _LastName = value;
                NotifyPropertyChanged("LastName");
                NotifyPropertyChanged("Salutation");
            }
        }
        public string Salutation
        {
            get { return _Salutation; }
            set
            {
                _Salutation = value;
                NotifyPropertyChanged("Salutation");
            }
        }
        public string Qualifications
        {
            get { return _Qualifications; }
            set
            {
                _Qualifications = value;
                NotifyPropertyChanged("Qualifications");
            }
        }
        public string Position
        {
            get { return _Position; }
            set
            {
                _Position = value;
                NotifyPropertyChanged("Position");
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
        public string Department
        {
            get { return _Department; }
            set
            {
                _Department = value;
                NotifyPropertyChanged("Department");
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
        [Email(ErrorMessage = "Invalid email address. eg. john@sample.com")]
        [XmlElement(IsNullable = true)]
        public string EmailAddress
        {
            get { return _EmailAddress; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                    ValidateProperty("EmailAddress", value);
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
        public string StreetAddress1
        {
            get { return _StreetAddress1; }
            set
            {
                _StreetAddress1 = value;
                NotifyPropertyChanged("StreetAddress1");
            }
        }
        public string StreetAddress2
        {
            get { return _StreetAddress2; }
            set
            {
                _StreetAddress2 = value;
                NotifyPropertyChanged("StreetAddress2");
            }
        }
        public string StreetCity
        {
            get { return _StreetCity; }
            set
            {
                _StreetCity = value;
                NotifyPropertyChanged("StreetCity");
            }
        }
        public string StreetState
        {
            get { return _StreetState; }
            set
            {
                _StreetState = value;
                NotifyPropertyChanged("StreetState");
            }
        }
        public string StreetPostalCode
        {
            get { return _StreetPostalCode; }
            set
            {
                _StreetPostalCode = value;
                NotifyPropertyChanged("StreetPostalCode");
            }
        }
        public string StreetCountry
        {
            get { return _StreetCountry; }
            set
            {
                _StreetCountry = value;
                NotifyPropertyChanged("StreetCountry");
            }
        }
        public string PostalAddress1
        {
            get { return _PostalAddress1; }
            set
            {
                _PostalAddress1 = value;
                NotifyPropertyChanged("PostalAddress1");
            }
        }
        public string PostalAddress2
        {
            get { return _PostalAddress2; }
            set
            {
                _PostalAddress2 = value;
                NotifyPropertyChanged("PostalAddress2");
            }
        }
        public string PostalCity
        {
            get { return _PostalCity; }
            set
            {
                _PostalCity = value;
                NotifyPropertyChanged("PostalCity");
            }
        }
        public string PostalState
        {
            get { return _PostalState; }
            set
            {
                _PostalState = value;
                NotifyPropertyChanged("PostalState");
            }
        }
        public string PostalPostalCode
        {
            get { return _PostalPostalCode; }
            set
            {
                _PostalPostalCode = value;
                NotifyPropertyChanged("PostalPostalCode");
            }
        }
        public string PostalCountry
        {
            get { return _PostalCountry; }
            set
            {
                _PostalCountry = value;
                NotifyPropertyChanged("PostalCountry");
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
        public string SignatureLine1
        {
            get { return _SignatureLine1; }
            set
            {
                _SignatureLine1 = value;
                NotifyPropertyChanged("SignatureLine1");
            }
        }
        public string SignatureLine2
        {
            get { return _SignatureLine2; }
            set
            {
                _SignatureLine2 = value;
                NotifyPropertyChanged("SignatureLine2");
            }
        }


        public List<string> States
        {
            get
            {
                if (_States == null)
                    _States = new List<string>() { "ACT", "NSW", "NT", "QLD", "SA", "TAS", "VIC", "WA" };
                return _States;
            }
            set { _States = value; }
        }
        public List<string> Titles
        {
            get
            {
                if (_Titles == null)
                    _Titles = new List<string>() { "Dr", "Miss", "Mr", "Mrs", "Ms", "Prof" };
                return _Titles;
            }
            set { _Titles = value; }
        }

        public bool IsDeliveryByEmail { get { return Delivery != null && Delivery.ToLower().Contains(Resources.Delivery_Email.ToLower()); } }
        public bool IsDeliveryByFax { get { return Delivery != null && Delivery.ToLower().Contains(Resources.Delivery_Facsimile.ToLower()); } }
        public bool IsDeliveryByOther { get { return Delivery != null && Delivery.ToLower().Contains(Resources.Delivery_Other.ToLower()); } }

        public Visibility EmailVisibility { get { return Delivery != null && Delivery.ToLower().Contains(Resources.Delivery_Email.ToLower()) ? Visibility.Visible : Visibility.Collapsed; } }
        public Visibility FaxVisibility { get { return Delivery != null && Delivery.ToLower().Contains(Resources.Delivery_Facsimile.ToLower()) ? Visibility.Visible : Visibility.Collapsed; } }
        public Visibility AddressVisibility { get { return EmailVisibility == Visibility.Collapsed && FaxVisibility == Visibility.Collapsed ? Visibility.Visible : Visibility.Collapsed; } }

        public bool IsEmpty()
        {
            if (!string.IsNullOrEmpty(Title))
                return false;
            if (!string.IsNullOrEmpty(Initials))
                return false;
            if (!string.IsNullOrEmpty(FullName))
                return false;
            if (!string.IsNullOrEmpty(FirstName))
                return false;
            if (!string.IsNullOrEmpty(LastName))
                return false;
            if (!string.IsNullOrEmpty(Salutation))
                return false;
            if (!string.IsNullOrEmpty(Qualifications))
                return false;
            if (!string.IsNullOrEmpty(Position))
                return false;
            if (!string.IsNullOrEmpty(CompanyName))
                return false;
            if (!string.IsNullOrEmpty(Department))
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
            if (!string.IsNullOrEmpty(_StreetAddress1))
                return false;
            if (!string.IsNullOrEmpty(_StreetAddress2))
                return false;
            if (!string.IsNullOrEmpty(_StreetCity))
                return false;
            if (!string.IsNullOrEmpty(_StreetState))
                return false;
            if (!string.IsNullOrEmpty(_StreetPostalCode))
                return false;
            if (!string.IsNullOrEmpty(_StreetCountry))
                return false;
            if (!string.IsNullOrEmpty(_PostalAddress1))
                return false;
            if (!string.IsNullOrEmpty(_PostalAddress2))
                return false;
            if (!string.IsNullOrEmpty(_PostalCity))
                return false;
            if (!string.IsNullOrEmpty(_PostalState))
                return false;
            if (!string.IsNullOrEmpty(_PostalPostalCode))
                return false;
            if (!string.IsNullOrEmpty(_PostalCountry))
                return false;
            if (!string.IsNullOrEmpty(_SignatureLine1))
                return false;
            if (!string.IsNullOrEmpty(_SignatureLine2))
                return false;
            return true;
        }

        public void Clear()
        {
            Title = string.Empty;
            Initials = string.Empty;
            FullName = string.Empty;
            FirstName = string.Empty;
            LastName = string.Empty;
            Salutation = string.Empty;
            Qualifications = string.Empty;
            Position = string.Empty;
            CompanyName = string.Empty;
            Department = string.Empty;
            PhoneNumber = string.Empty;
            FaxNumber = string.Empty;
            EmailAddress = string.Empty;
            Address = string.Empty;
            StreetAddress1 = string.Empty;
            StreetAddress2 = string.Empty;
            StreetCity = string.Empty;
            StreetState = string.Empty;
            StreetPostalCode = string.Empty;
            StreetCountry = string.Empty;
            PostalAddress1 = string.Empty;
            PostalAddress2 = string.Empty;
            PostalCity = string.Empty;
            PostalState = string.Empty;
            PostalPostalCode = string.Empty;
            PostalCountry = string.Empty;
            Country = string.Empty;
            Delivery = string.Empty;
            SignatureLine1 = string.Empty;
            SignatureLine2 = string.Empty;
        }
        public void Copy(Contact from)
        {
            Title = from.Title;
            Initials = from.Initials;
            FullName = from.FullName;
            FirstName = from.FirstName;
            LastName = from.LastName;
            Salutation = from.Salutation;
            Qualifications = from.Qualifications;
            Position = from.Position;
            CompanyName = from.CompanyName;
            Department = from.Department;
            PhoneNumber = from.PhoneNumber;
            FaxNumber = from.FaxNumber;
            EmailAddress = from.EmailAddress;
            Address = from.Address;
            StreetAddress1 = from.StreetAddress1;
            StreetAddress2 = from.StreetAddress2;
            StreetCity = from.StreetCity;
            StreetState = from.StreetState;
            StreetPostalCode = from.StreetPostalCode;
            StreetCountry = from.StreetCountry;
            PostalAddress1 = from.PostalAddress1;
            PostalAddress2 = from.PostalAddress2;
            PostalCity = from.PostalCity;
            PostalState = from.PostalState;
            PostalPostalCode = from.PostalPostalCode;
            PostalCountry = from.PostalCountry;
            Country = from.Country;
            Delivery = from.Delivery;
            SignatureLine1 = from.SignatureLine1;
            SignatureLine2 = from.SignatureLine2;
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

        public override string ToString() { return FullName; }
    }
}
