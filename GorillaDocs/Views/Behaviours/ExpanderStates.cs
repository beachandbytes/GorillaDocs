using GorillaDocs.Properties;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Xml.Serialization;

namespace GorillaDocs.Views
{
    [Serializable]
    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    [XmlRoot("ExpanderStates")]
    public class ExpanderStates : List<ExpanderState>
    {
        const string IsExpanded = " IsExpanded";
        public ExpanderStates() : base() { }
        public ExpanderStates(IEnumerable<ExpanderState> collection) : base(collection) { }
        public static ExpanderStates Get()
        {
            var states = Settings.Default.Category_Expander_States;
            return states == null ? new ExpanderStates() : states;
        }
        public void Set(string GroupName, bool Expanded)
        {
            if (this.Any(x => x.GroupName == GroupName + IsExpanded))
                this.First(x => x.GroupName == GroupName + IsExpanded).IsExpanded = Expanded;
            else
                Add(new ExpanderState() { GroupName = GroupName + IsExpanded, IsExpanded = Expanded });
            Settings.Default.Category_Expander_States = this;
            Settings.Default.Save();
        }
        public bool Get(string GroupName)
        {
            var state = this.FirstOrDefault(x => x.GroupName == GroupName + IsExpanded);
            return state == null ? true : state.IsExpanded;
        }
    }
}
