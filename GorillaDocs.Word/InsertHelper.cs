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

        public static Wd.Range InsertFromFile(this Wd.Range range, string Path, string BookmarkName = null)
        {
            if (range.IsCollapsed() && range.InContentControlOrContainsControls())
                range.MoveOutOfContentControl();
            if ((bool)range.Information[Wd.WdInformation.wdWithInTable])
                range.MoveOutOfTable();

            if (string.IsNullOrEmpty(BookmarkName))
                range.InsertFile(Path);
            else
            {
                ((Wd.Document)range.Parent).Bookmarks.Delete(BookmarkName);
                range.InsertFile(Path, BookmarkName);
                range = range.Bookmarks[BookmarkName].Range;
                range.Bookmarks.Delete(BookmarkName);
            }
            return range;
        }

        /// <summary>
        /// Prevents strange errors when the range is near a Content Control
        /// </summary>
        public static Wd.Range InsertFile_Safe(this Wd.Range range, string FileName, string Bookmark = "")
        {
            ((Wd.Document)range.Parent).Bookmarks.DeleteIfExists(Bookmark);
            range.Text = ".";
            range.Delete();
            range.InsertFile(FileName, Bookmark);
            range = range.Bookmarks[Bookmark].Range;
            range.Bookmarks.Delete(Bookmark);
            return range;
        }
       
        /// <summary>
        /// Inserts a file at the given range.
        /// Note: 
        /// If you want to insert Headers and Footers with the file, then a trailing section break is required.
        /// If you want to different margins or page layout, then a leading section break is required.
        /// When inserting AsNewSection a bookmark MUST be used to avoid trouble with the last paragraph mark in the document.
        /// </summary>
        /// <param name="range"></param>
        /// <param name="Path"></param>
        /// <returns></returns>
        public static Wd.Range InsertSectionFromFile(this Wd.Range range, string Path)
        {
            if (range.InContentControlOrContainsControls())
                range.MoveOutOfContentControl(Wd.WdCollapseDirection.wdCollapseStart);
            if ((bool)range.Information[Wd.WdInformation.wdWithInTable])
                range.MoveOutOfTable();

            if (range.IsStartOfSection())
                range = range.InsertSectionAtStartOfSection(Path);
            else if (range.IsEndOfDocument())
                range = range.InsertSectionAtEndOfDocument(Path);
            else if (range.IsEndOfSection())
                range = range.InsertSectionAtEndOfSection(Path);
            else
                range = range.InsertSectionInMiddldOfSection(Path);

            if (range.Document.Bookmarks.Exists("Section"))
            {
                range = range.Document.Bookmarks["Section"].Range;
                range.Document.Bookmarks["Section"].Delete();
            }

            return range;
        }

        static Wd.Range InsertSectionAtStartOfSection(this Wd.Range range, string Path)
        {
            range.InsertFile_Safe(Path, "Section");
            range.Sections[1].Delete();
            return range;
        }

        static Wd.Range InsertSectionInMiddldOfSection(this Wd.Range range, string Path)
        {
            var AddedSection = range.AddSection();
            range.InsertFile(Path, "Section");
            AddedSection.Next().Delete();
            return range;
        }

        static Wd.Range InsertSectionAtEndOfSection(this Wd.Range range, string Path)
        {
            var NextSection = range.Sections[1].Next();
            var AddedSection = range.AddSection();
            range.InsertFile(Path, "Section");
            AddedSection.Next().Delete();
            NextSection.Previous().Delete();
            return range;
        }

        static Wd.Range InsertSectionAtEndOfDocument(this Wd.Range range, string Path)
        {
            var AddedSection = range.AddSection();
            range.InsertFile(Path, "Section");
            AddedSection.Next().Delete();
            range.Document.Sections.Last.Delete();
            return range;
        }

    }
}
