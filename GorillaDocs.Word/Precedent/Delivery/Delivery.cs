using System;
using System.Linq;
using GDP = GorillaDocs.Properties;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.Precedent.Delivery
{
    public abstract class Delivery1
    {
        protected readonly Wd.Document doc;
        protected readonly Wd.ContentControl control;

        public Delivery1(Wd.ContentControl control)
        {
            doc = control.Range.Document;
            this.control = control;
        }

        public abstract void Update();

        protected bool IsEmail { get { return control.Range.Text.Contains(GDP.Resources.Delivery_Email); } }
        protected bool IsFacsimile { get { return control.Range.Text.Contains(GDP.Resources.Delivery_Facsimile); } }

        protected virtual Wd.ContentControl DetailsControl
        {
            get
            {
                var range = control.Range;
                range.MoveEndUntil("\r\v\a");
                return range.ContentControls(x => x.Title == control.Title + " Details").FirstOrDefault();
            }
        }

        protected void UpdateEmail()
        {
            if (DetailsControl == null)
                AddDetailsControl();
            SetMapping("EmailAddress");
            //ConvertToHyperlink();
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
            //TODO: Can not use a RichText control until figure out how to strip the Hyperlink XML markup from CustomXMLPart before loading it into the Edit Details dialog.
            Wd.ContentControl detailsControl;
            //if (doc.Application.Version.StartsWith("14"))
            //{
                detailsControl = range.ContentControls.Add_Safely(Wd.WdContentControlType.wdContentControlText);
                detailsControl.MultiLine = true;
            //}
            //else
            //    detailsControl = range.ContentControls.Add_Safely(Wd.WdContentControlType.wdContentControlRichText);
            detailsControl.Title = control.Title + " Details";
            return detailsControl;
        }

        protected void SetMapping(string value)
        {
            var xpath = control.XMLMapping.XPath;
            DetailsControl.XMLMapping.SetMapping(xpath.Substring(0, xpath.LastIndexOf(':')) + String.Format(":{0}[1]", value));
        }
        void ConvertToHyperlink() { DetailsControl.Range.AutoFormat(); }
    }
}
