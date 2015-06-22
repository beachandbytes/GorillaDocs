using GorillaDocs.libs.PostSharp;
using System;
using System.ComponentModel;

namespace GorillaDocs.Models
{
    [Log]
    public class Party : INotifyPropertyChanged, IEquatable<Party>
    {
        string _Name;
        string _RegistrationNumber;
        string _RegisteredAddress;
        string _DefinedName;

        public string Name
        {
            get { return _Name; }
            set
            {
                _Name = value;
                NotifyPropertyChanged("Name");
            }
        }
        public string RegistrationNumber
        {
            get { return _RegistrationNumber; }
            set
            {
                _RegistrationNumber = value;
                NotifyPropertyChanged("RegistrationNumber");
            }
        }
        public string RegisteredAddress
        {
            get { return _RegisteredAddress; }
            set
            {
                _RegisteredAddress = value;
                NotifyPropertyChanged("RegisteredAddress");
            }
        }
        public string DefinedName
        {
            get { return _DefinedName; }
            set
            {
                _DefinedName = value;
                NotifyPropertyChanged("DefinedName");
            }
        }

        public bool IsEmpty()
        {
            if (!string.IsNullOrEmpty(Name))
                return false;
            if (!string.IsNullOrEmpty(RegistrationNumber))
                return false;
            if (!string.IsNullOrEmpty(RegisteredAddress))
                return false;
            if (!string.IsNullOrEmpty(DefinedName))
                return false;
            return true;
        }
        public void Clear()
        {
            Name = string.Empty;
            RegistrationNumber = string.Empty;
            RegisteredAddress = string.Empty;
            DefinedName = string.Empty;
            Notify();
        }
        void Notify()
        {
            NotifyPropertyChanged("FullEntityName");
            NotifyPropertyChanged("RegistrationNumber");
            NotifyPropertyChanged("RegisteredAddress");
            NotifyPropertyChanged("DefinedName");
        }

        public bool Equals(Party other)
        {
            if (other == null)
                return false;
            if (other == this)
                return true;
            return false;
        }
        public override bool Equals(object obj) { return Equals(obj as Party); }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void NotifyPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public static class PartyHelper
    {
        public static bool IsNullOrEmpty(this Party party)
        {
            return party == null || party.IsEmpty();
        }
    }
}
