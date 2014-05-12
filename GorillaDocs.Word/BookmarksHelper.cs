using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class BookmarksHelper
    {
        public static void Delete(this Wd.Bookmarks bookmarks, string BookmarkName)
        {
            if (bookmarks.Exists(BookmarkName))
                bookmarks[BookmarkName].Delete();
        }

        public static Wd.Range Replace(this Wd.Bookmarks bookmarks, string BookmarkName, string value = null, bool restoreBookmark = false, bool trim = true, bool deleteLine = false, bool deleteRow = false)
        {
            if (value == null) value = string.Empty;
            if (value == "-None-") value = "";
            if (trim) value = value.Trim();
            if (bookmarks.Exists(BookmarkName))
            {
                Wd.Range range = bookmarks[BookmarkName].Range;
                if (string.IsNullOrEmpty(value))
                    range.Delete();
                else
                    range.Text = value;

                if (string.IsNullOrEmpty(value) && deleteLine)
                    range.DeleteParagraphFromRange();
                else if (string.IsNullOrEmpty(value) && deleteRow)
                {
                    range.Expand(Wd.WdUnits.wdRow);
                    if (!string.IsNullOrEmpty(range.Text))
                        range.Cut();
                    ClipboardHelper.Clear();
                }

                if (restoreBookmark)
                    range.Bookmarks.Add(BookmarkName);
                else if (bookmarks.Exists(BookmarkName))
                    bookmarks[BookmarkName].Delete();
                return range;
            }
            else
                return null;
        }


    }
}
