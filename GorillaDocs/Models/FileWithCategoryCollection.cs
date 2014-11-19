using System;
using System.Collections.Generic;
using System.Configuration;
using System.Xml.Serialization;

namespace GorillaDocs.Models
{
    [Serializable]
    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    [XmlRoot("FileWithCategoryCollection")]
    public class FileWithCategoryCollection : List<FileWithCategory>
    {
        public FileWithCategoryCollection() : base() { }
        public FileWithCategoryCollection(IEnumerable<FileWithCategory> collection) : base(collection) { }
    }
}
