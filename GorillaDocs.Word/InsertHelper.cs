using GorillaDocs.libs.PostSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    [Log]
    public static class InsertHelper
    {
        public static void InsertFromTemplate(this Wd.Bookmarks bookmarks, string BookmarkName, bool required = false)
        {
            if (bookmarks.Exists(BookmarkName))
            {
                var range = bookmarks[BookmarkName].Range;
                var doc = (Wd.Document)bookmarks.Parent;
                var template = (Wd.Template)doc.get_AttachedTemplate();
                range.InsertFile(template.FullName, BookmarkName);
            }
            else
                if (required)
                    throw new InvalidOperationException(string.Format("Unable to find the required bookmark '{0}'.", BookmarkName));
        }

        public static Wd.Range InsertFromTemplate(this Wd.Range range, string BookmarkName)
        {
            if (range.IsCollapsed() && range.InContentControlOrContainsControls())
                range.MoveOutOfContentControl();

            ((Wd.Document)range.Parent).Bookmarks.Delete(BookmarkName);
            var template = (Wd.Template)range.Document.get_AttachedTemplate();
            range.InsertFile(template.FullName, BookmarkName);
            range = range.Bookmarks[BookmarkName].Range;
            range.Bookmarks.Delete(BookmarkName);
            return range;
        }

        public static Wd.Range InsertFromFile(this Wd.Bookmarks bookmarks, string Path, string BookmarkName)
        {
            if (bookmarks.Exists(BookmarkName))
            {
                var range = bookmarks[BookmarkName].Range;
                ((Wd.Document)range.Parent).Bookmarks.Delete(BookmarkName);
                range.InsertFile(Path, BookmarkName);
                range = range.Bookmarks[BookmarkName].Range;
                return range;
            }
            return null;
        }

        public static Wd.Range InsertFromFile(this Wd.Range range, string Path, string BookmarkName)
        {
            if (range.IsCollapsed() && range.InContentControlOrContainsControls())
                range.MoveOutOfContentControl();

            ((Wd.Document)range.Parent).Bookmarks.Delete(BookmarkName);
            range.InsertFile(Path, BookmarkName);
            range = range.Bookmarks[BookmarkName].Range;
            range.Bookmarks.Delete(BookmarkName);
            return range;
        }
    }
}
