using GorillaDocs.libs.PostSharp;
using GorillaDocs.Models;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;

namespace GorillaDocs.ViewModels
{
    [Log]
    public class ChangeOfficeViewModel : BaseViewModel
    {
        readonly IUserSettings settings;
        public ChangeOfficeViewModel(IUserSettings settings)
        {
            this.settings = settings;
            OKCommand = new RelayCommand(OKPressed, CanPressOK);
        }

        List<IOffice> offices = new List<IOffice>();
        public virtual List<IOffice> Offices { get { return offices; } set { offices = value; } }

        IOffice office;
        public virtual IOffice Office
        {
            get
            {
                if (office == null)
                    office = Offices.LastSelected(settings);
                if (office == null && Offices.Count > 0)
                    office = Offices[0];
                return office;
            }
            set
            {
                settings.Last_Selected_Office = value.Name;
                this.office = value;
                NotifyPropertyChanged("Office");
            }
        }

        public ICommand OKCommand { get; set; }
        public virtual bool CanPressOK() { return true; }
        public virtual void OKPressed() { DialogResult = true; }
    }
}
