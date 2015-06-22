using GorillaDocs.Models;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public class CCs
    {
        public delegate void CcEventHandler(int index);

        readonly Wd.Document doc;
        readonly string[] controlTags = null;

        public CCs(Wd.Document doc, string[] controlTags)
        {
            this.doc = doc;
            this.controlTags = controlTags;
        }

        public void Delete()
        {
            var control = doc.ContentControls(x => !string.IsNullOrEmpty(x.Tag) && x.Tag.StartsWith("CC")).FirstOrDefault();
            if (control != null)
                control.Range.Tables[1].Delete();
        }

        public Wd.Range UpdateControls(IList<Contact> contacts)
        {
            try
            {
                var control = doc.ContentControls(x => x.Tag == controlTags.First()).First();
                var range = control.Range.Cells[1].Range;
                range.MoveEnd(Wd.WdUnits.wdCharacter, -1);
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
                range.Cells[1].Range.Characters.Last.Previous().Delete();
                return range.Tables[1].Range.MoveOutOfTable();
            }
            finally
            {
                ClipboardHelper.Clear();
            }
        }

        void UpdateContentControlMappings(Wd.Range range, int i)
        {
            foreach (Wd.ContentControl control in range.ContentControls)
                if (!control.XMLMapping.SetMapping(UpdateContactIndex(control.XMLMapping.XPath, i)))
                {
                    Message.LogWarning("Unable to set Mapping '{0}' for control '{1}'. This may be because no data was entered by user.", control.XMLMapping.XPath, control.Title);
                    control.Delete(true);
                }
        }

        static string UpdateContactIndex(string xPath, int contactIndex) { return Regex.Replace(xPath, @"Contact\[[\d]*\]", string.Format("Contact[{0}]", contactIndex)); }
    }
}
