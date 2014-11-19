using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Xml.Serialization;

namespace GorillaDocs.Models
{
    [Serializable]
    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    [XmlRoot("RecentFilesCollection")]
    public class RecentFilesCollection
    {
        public List<string> keys = new List<string>();
        public List<List<FileWithCategory>> values = new List<List<FileWithCategory>>();

        public bool Contains(string Key) { return keys.Contains(Key); }
        public List<FileWithCategory> FirstOrDefault(string Key)
        {
            if (Contains(Key))
                return values[keys.FindIndex(x => x == Key)];
            else
                return new List<FileWithCategory>();
        }
        public void Add(string Key, List<FileWithCategory> Values)
        {
            int index = keys.FindIndex(x => x == Key);
            if (index > -1)
            {
                keys.RemoveAt(index);
                values.RemoveAt(index);
            }

            keys.Add(Key);
            values.Add(Values);
        }
    }
}
