using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.ControlManagers
{
    public class ControlManager_EndDocument
    {
        readonly Wd.Document Doc;
        readonly Dictionary<string, bool?> controls = new Dictionary<string, bool?>();

        public ControlManager_EndDocument(Wd.Document Doc) { this.Doc = Doc; }

        public void Add(string ControlTitle, bool? Replace = null) { controls.Add(ControlTitle, Replace); }

        public void ReplaceMissingControls()
        {
            Doc.EnsureTheLastParagraphIsEmpty();
            var range = Doc.Range().CollapseEnd();
            foreach (KeyValuePair<string, bool?> control in controls)
                ReplaceControl(control, ref range);
        }

        void ReplaceControl(KeyValuePair<string, bool?> control, ref Wd.Range range)
        {
            var bookmark = "EditDetails_" + control.Key.Replace(' ', '_');
            var controls = Doc.ContentControls(x => x.Title == control.Key);

            if (control.Value == null)
            {
                if (ControlsFound(controls))
                    controls.DeleteRowsOrParagraphs();
                range = range.InsertFile_Safe1(Doc.Template().FullName, bookmark).CollapseStart();
            }
            else if (control.Value == true)
            {
                if (ControlsFound(controls))
                    range = controls[0].CollapsePastRowOrParagraph(Wd.WdCollapseDirection.wdCollapseStart);
                else
                    range = range.InsertFile_Safe1(Doc.Template().FullName, bookmark).CollapseStart();
            }
            else if (control.Value == false && controls.Count != 0)
                range = controls[0].CollapsePastRowOrParagraph(Wd.WdCollapseDirection.wdCollapseStart);

            Doc.Bookmarks.DeleteIfExists(bookmark);
        }

        static bool ControlsFound(IList<Wd.ContentControl> controls) { return controls.Count != 0; }
    }
}
