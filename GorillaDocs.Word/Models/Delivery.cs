using GorillaDocs;
using System;
using System.Collections.Generic;
using System.Linq;
using GDP = GorillaDocs.Properties;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.Models
{
    public abstract class Delivery
    {
        protected readonly Wd.Document doc;
        protected readonly Wd.ContentControl control;

        public Delivery(Wd.ContentControl control)
        {
            doc = (Wd.Document)control.Parent;
            this.control = control;
        }

        public abstract void Update();

        protected bool IsEmail { get { return control.Range.Text.Contains(GDP.Resources.Delivery_Email); } }
        protected bool IsFacsimile { get { return control.Range.Text.Contains(GDP.Resources.Delivery_Facsimile); } }

        protected virtual Wd.ContentControl DetailsControl
        {
            get
            {
                string tag = control.Tag.Substring(0, control.Tag.IndexOf(GDP.Resources.ContentControl_DeliveryEnding)) + "Details";
                return doc.ContentControls.Find(tag);
            }
        }

        protected void UpdateEmail()
        {
            if (DetailsControl == null)
                AddDetailsControl();
            SetMapping("EmailAddress");
        }
        protected void UpdateFacsimile()
        {
            if (DetailsControl == null)
                AddDetailsControl();
            SetMapping("FaxNumber");
        }

        protected abstract void AddDetailsControl();

        protected Wd.ContentControl AddDetailsControl(Wd.Range range)
        {
            var detailsControl = range.ContentControls.Add_Safely(Wd.WdContentControlType.wdContentControlText);
            detailsControl.Title = control.Title + " Details";
            detailsControl.Tag = control.Tag.Remove(" Delivery") + " Details";
            return detailsControl;
        }

        protected void SetMapping(string value)
        {
            var xpath = control.XMLMapping.XPath;
            DetailsControl.XMLMapping.SetMapping(xpath.Substring(0, xpath.LastIndexOf(':')) + String.Format(":{0}[1]", value));
        }
    }
}
