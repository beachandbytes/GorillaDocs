using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.ControlManagers
{
    public class ControlManager_Table
    {
        readonly Wd.Document Doc;
        readonly string TableDesc;
        readonly Dictionary<string, bool?> controls = new Dictionary<string, bool?>();

        public ControlManager_Table(Wd.Document Doc, string TableDesc)
        {
            this.Doc = Doc;
            this.TableDesc = TableDesc;
        }

        public void Add(string ControlTitle, bool? Replace = null) { controls.Add(ControlTitle, Replace); }

        public void ReplaceMissingControls()
        {
            var table = Doc.Tables(x => x.Descr == TableDesc).FirstOrDefault();
            if (table == null)
                ReplaceEntireTable();
            else
                ReplaceRows(table);
        }

        void ReplaceEntireTable()
        {
            var bookmark = "EditDetails_" + TableDesc.Replace(' ', '_');

            throw new NotImplementedException();
        }

        void ReplaceRows(Wd.Table table)
        {
            var range = table.Range.CollapseStart();
            foreach (KeyValuePair<string, bool?> control in controls)
                ReplaceRow(control, ref range);
        }

        void ReplaceRow(KeyValuePair<string, bool?> control, ref Wd.Range range)
        {
            var bookmark = "EditDetails_" + control.Key.Replace(' ', '_');
            var controls = Doc.ContentControls(x => x.Title == control.Key);

            if (control.Value == null)
            {
                if (ControlsFound(controls))
                    range.Rows[1].Delete();
                range.InsertFile_Safe1(Doc.Template().FullName, bookmark);
            }
            else if (control.Value == true)
                if (!(ControlsFound(controls)))
                    range.InsertFile_Safe1(Doc.Template().FullName, bookmark);

            if (range.Information[Wd.WdInformation.wdWithInTable])
                range = range.Rows[1].Range.CollapseEnd();
        }

        static bool ControlsFound(IList<Wd.ContentControl> controls) { return controls.Count != 0; }
    }
}
