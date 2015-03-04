using GorillaDocs.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs.ViewModels
{
    public class SelectTemplateViewModel : ChangeOfficeViewModel
    {
        public delegate void SelectTemplateEventHandler(string path);

        const int RECENT_TEMPLATES = 1;
        const int ALL_TEMPLATES = 0;
        readonly MicrosoftApplication app;
        readonly IUserSettings settings;

        public SelectTemplateViewModel(IUserSettings settings, MicrosoftApplication app)
            : base(settings)
        {
            this.app = app;
            this.settings = settings;
        }

        public List<FileWithCategory> Templates { get; set; }
        public List<FileWithCategory> RecentTemplates { get; set; }

        public override List<IOffice> Offices
        {
            get { return base.Offices; }
            set
            {
                base.Offices = value;
                RefreshTemplates();
            }
        }

        public override IOffice Office
        {
            get { return base.Office; }
            set
            {
                settings.Last_Selected_Office = value.Name;
                base.Office = value;
                RefreshTemplates();
            }
        }

        void RefreshTemplates()
        {
            if (app != null)
                Templates = Office.GetTemplates(app);
            NotifyPropertyChanged("Templates");

            RecentTemplates = Office.RecentFiles;
            NotifyPropertyChanged("RecentTemplates");
        }

        public int SelectedTab
        {
            get { return settings.Last_Selected_Templates_Tab; }
            set
            {
                settings.Last_Selected_Templates_Tab = value;
                NotifyPropertyChanged("SelectedCategory");
            }
        }

        public FileWithCategory SelectedTemplate
        {
            get
            {
                if (Templates == null)
                    return null;
                else
                {
                    FileWithCategory file = Templates.Where(x => x.NameWithoutExtension == settings.Last_Selected_Template).FirstOrDefault();
                    return file ?? Templates.FirstOrDefault();
                }
            }
            set
            {
                try
                {
                    if (value != null)
                        settings.Last_Selected_Template = value.NameWithoutExtension;
                }
                catch (Exception ex)
                {
                    Message.LogError(ex);
                }
            }
        }
        FileWithCategory _SelectedRecentTemplate;
        public FileWithCategory SelectedRecentTemplate
        {
            get
            {
                if (_SelectedRecentTemplate == null && RecentTemplates != null && RecentTemplates.Count > 0)
                    _SelectedRecentTemplate = RecentTemplates[0];
                return _SelectedRecentTemplate;
            }
            set { _SelectedRecentTemplate = value; }
        }
        public FileWithCategory GetTemplate()
        {
            if (SelectedTab == ALL_TEMPLATES)
                return SelectedTemplate;
            else
                return SelectedRecentTemplate;
        }

        public override bool CanPressOK()
        {
            switch (SelectedTab)
            {
                case ALL_TEMPLATES:
                    return SelectedTemplate != null;
                case RECENT_TEMPLATES:
                    return SelectedRecentTemplate != null;
                default:
                    throw new InvalidOperationException("You must select a Tab");
            }
        }
        public override void OKPressed()
        {
            try
            {
                DialogResult = true;
                UpdateRecentFiles();
            }
            catch (Exception ex)
            {
                Message.LogError(ex);
            }
        }

        void UpdateRecentFiles()
        {
            RecentTemplates.RemoveAll(x => x.NameWithoutExtension == GetTemplate().NameWithoutExtension);
            RecentTemplates.Insert(0, GetTemplate());
            Office.RecentFiles = RecentTemplates;
        }
    }
}
