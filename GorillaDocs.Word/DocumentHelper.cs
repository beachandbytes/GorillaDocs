using GorillaDocs.libs.PostSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    [Log]
    public static class DocumentHelper
    {
        public static bool ActiveDocumentExists(this Wd.Application application)
        {
            if (application.Documents.Count == 0)
                return false;
            else if (application.ActiveProtectedViewWindow != null)
                return false;
            return true;
        }

        public static void ReplaceTableWithBookmarkFromTemplate(this Wd.Document doc, string TableDescription, string TemplateBookmark, Wd.WdCollapseDirection CollapseDirectionIfTableMissing = Wd.WdCollapseDirection.wdCollapseStart)
        {
            Wd.Range range = doc.FindTableRange(TableDescription);
            if (range == null)
                if (CollapseDirectionIfTableMissing == Wd.WdCollapseDirection.wdCollapseStart)
                    range = doc.Range().CollapseStart();
                else
                    range = doc.Range().CollapseEnd();
            else
            {
                range.ContentControls.Unlock();
                range.Rows.Delete();
            }
            range.InsertFromTemplate(TemplateBookmark);
        }

        public static Wd.Range FindTableRange(this Wd.Document doc, string description)
        {
            Wd.Table table = doc.Range().FindTable(description);
            if (table == null)
                return null;
            else
                return table.Range;
        }

        public static string NameWithoutExtension(this Wd.Document doc)
        {
            return doc.Name.Substring(0, doc.Name.LastIndexOf('.'));
        }
    }
}
