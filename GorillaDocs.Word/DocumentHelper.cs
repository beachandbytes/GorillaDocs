using GorillaDocs.libs.PostSharp;
using System;
using System.IO;
using O = Microsoft.Office.Core;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
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
            return doc.Name.Contains(".") ? doc.Name.Substring(0, doc.Name.LastIndexOf('.')) : doc.Name;
        }

        [System.Diagnostics.DebuggerStepThrough]
        public static bool IsTemplate(this Wd.Document doc)
        {
            return doc.FullName == doc.get_AttachedTemplate().Fullname;
        }

        public static O.CustomXMLPart CustomXmlPart(this Wd.Document Doc, string Namespace)
        {
            var parts = Doc.CustomXMLParts.SelectByNamespace(Namespace);
            if (parts.Count == 0)
                throw new InvalidOperationException(string.Format("Unable to find CustomXmlPart with Namespace '{0}'", Namespace));
            if (parts.Count > 1)
                throw new InvalidOperationException(string.Format("There are more than 1 CustomXmlPart with Namespace '{0}'", Namespace));
            return parts[1];
        }

        public static bool Exists(this Wd.Documents documents, FileInfo file)
        {
            foreach (Wd.Document doc in documents)
                if (doc.FullName == file.FullName)
                    return true;
            return false;
        }

        public static Wd.Document First(this Wd.Documents documents, FileInfo file)
        {
            foreach (Wd.Document doc in documents)
                if (doc.FullName == file.FullName)
                    return doc;
            return null;
        }

        public static FileInfo Template(this Wd.Document doc)
        {
            Wd.Template template = doc.get_AttachedTemplate();
            return new FileInfo(template.FullName);
        }
    }
}
