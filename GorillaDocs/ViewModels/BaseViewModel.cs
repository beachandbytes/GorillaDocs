using GorillaDocs.libs.PostSharp;
using System.Xml.Serialization;

namespace GorillaDocs.ViewModels
{
    [Log]
    public class BaseViewModel : Notify
    {
        public BaseViewModel() { }

        bool? dialogResult;
        [XmlIgnore]
        public bool? DialogResult
        {
            get { return this.dialogResult; }
            set
            {
                this.dialogResult = value;
                NotifyPropertyChanged("DialogResult");
            }
        }
    }
}
