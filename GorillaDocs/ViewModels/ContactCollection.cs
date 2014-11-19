using GorillaDocs.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Linq;
using System.Xml.Serialization;

namespace GorillaDocs.ViewModels
{
    [Serializable]
    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    [XmlRoot("ContactCollection")]
    public class ContactCollection : ObservableCollection<Contact>
    {
        public ContactCollection() : base() { }
        public ContactCollection(IEnumerable<Contact> collection) : base(collection) { }
    }
}
