using GorillaDocs.Word;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Wd = Microsoft.Office.Interop.Word;
using GorillaDocs;
using GorillaDocs.Models;

namespace GorillaDocs.Word
{
    public class RepeatingContacts
    {
        readonly Wd.Document doc;
        readonly string[] controlTags = null;
        IList<string> Bookmarks
        {
            get
            {
                var bookmarks = new List<string>();
                foreach (string tag in controlTags)
                    bookmarks.Add("EditDetails_" + tag.Replace(' ', '_'));
                return bookmarks;
            }
        }
        readonly string deliveryTag = null;

        public RepeatingContacts(Wd.Document doc, string[] controlTags, string deliveryTag)
        {
            this.doc = doc;
            this.controlTags = controlTags;
            this.deliveryTag = deliveryTag;
        }

        public void Delete()
        {
            foreach (Wd.ContentControl control in doc.ContentControls(x => x.Tag == "Recipients"))
                control.Delete(true);
        }

        public Wd.Range UpdateControls(IList<Contact> contacts, Wd.WdCollapseDirection CollapseDirection = Wd.WdCollapseDirection.wdCollapseEnd)
        {
            try
            {
                var control = doc.ContentControls(x => x.Tag == controlTags.First()).First();
                var range = control.ParentContentControl.Range;
                range.MoveEnd(Wd.WdUnits.wdParagraph, 1);
                range.Copy();

                range.ContentControls.DeleteEmpty();
                range.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);
                for (int i = 2; i <= contacts.Count; i++)
                {
                    range.MoveOutOfContentControl();
                    range.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);
                    range.Paste();
                    UpdateContentControlMappings(range, i);
                    range.ContentControls.DeleteEmpty();
                }
                range = range.CollapseEnd();
                if (range.Paragraphs[1].IsEmpty())
                    range.Paragraphs[1].Range.Delete();

                doc.Bookmarks.DeleteIfExists(Bookmarks);
                UpdateDeliveryDetails();

                range.Move(Wd.WdUnits.wdCharacter, 1);
                return range;
            }
            finally
            {
                ClipboardHelper.Clear();
            }
        }

        void UpdateDeliveryDetails()
        {
            var controls = doc.ContentControls(x => x.Tag == deliveryTag);
            foreach (Wd.ContentControl control in controls)
                if (control.IsDelivery())
                    control.AsDelivery().Update();
        }

        void UpdateContentControlMappings(Wd.Range range, int i)
        {
            foreach (Wd.ContentControl control in range.ContentControls)
                if (control.XMLMapping != null && !string.IsNullOrEmpty(control.XMLMapping.XPath))
                    if (!control.XMLMapping.SetMapping(UpdateContactIndex(control.XMLMapping.XPath, i)))
                    {
                        Message.LogWarning("Unable to set Mapping '{0}' for control '{1}'. This may be because no data was entered by user.", control.XMLMapping.XPath, control.Title);
                        control.Delete(true);
                    }
        }

        static int GetContactIndex(string Value)
        {
            var expression = new Regex(@"Contact\[(\d+)\]");
            var matches = expression.Match(Value);
            return int.Parse(matches.Groups[1].Value);
        }
        static string UpdateContactIndex(string xPath, int contactIndex) { return Regex.Replace(xPath, @"Contact\[[\d]*\]", string.Format("Contact[{0}]", contactIndex)); }
    }
}
