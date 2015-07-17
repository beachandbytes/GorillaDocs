using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.ControlManagers
{
    public class ControlManager_StartDocument
    {
        readonly Wd.Document Doc;
        readonly List<ControlFragment> Fragments = new List<ControlFragment>();

        public ControlManager_StartDocument(Wd.Document Doc) { this.Doc = Doc; }

        public void Add(string ControlTitle, bool? Replace = null, string[] AdditionalControlTitles = null) { Fragments.Add(new ControlFragment() { Title = ControlTitle, Replace = Replace, AdditionalControlsTitles = AdditionalControlTitles }); }

        public void ReplaceMissingControls()
        {
            var range = Doc.Range().CollapseStart();
            foreach (ControlFragment fragment in Fragments)
                ReplaceControl(fragment, ref range);
        }

        void ReplaceControl(ControlFragment fragment, ref Wd.Range range)
        {
            var bookmark = "EditDetails_" + fragment.Title.Replace(' ', '_');
            var controls = Doc.ContentControls(x => x.Title == fragment.Title ||
                (fragment.AdditionalControlsTitles != null && Array.IndexOf(fragment.AdditionalControlsTitles, x.Title) >= 0));

            if (fragment.Replace == null)
            {
                if (ControlsFound(controls))
                    controls.DelteRowOrParagraph();
                range = range.InsertFile_Safe1(Doc.Template().FullName, bookmark).CollapseEnd();
            }
            else if (fragment.Replace == true)
            {
                if (ControlsFound(controls))
                    range = controls[0].CollapsePastRowOrParagraph(Wd.WdCollapseDirection.wdCollapseEnd);
                else
                    range = range.InsertFile_Safe1(Doc.Template().FullName, bookmark).CollapseEnd();
            }
            else if (fragment.Replace == false && controls.Count != 0)
                range = controls[0].CollapsePastRowOrParagraph(Wd.WdCollapseDirection.wdCollapseEnd);

            Doc.Bookmarks.DeleteIfExists(bookmark);
        }

        static bool ControlsFound(IList<Wd.ContentControl> controls) { return controls.Count != 0; }
    }
}