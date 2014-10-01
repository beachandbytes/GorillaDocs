using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class SearchHelper
    {
        public static Wd.Range Find(this Wd.Range range, string value)
        {
            var ShowAll = range.Document.ActiveWindow.View.ShowAll;

            try
            {
                range.Document.ActiveWindow.View.ShowAll = true;

                Wd.Range result = range.Duplicate;
                Wd.Find find = result.Find;
                find.ClearFormatting();
                find.Text = value;
                find.Forward = true;
                find.Wrap = Wd.WdFindWrap.wdFindStop;
                find.Forward = true;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                find.MatchWildcards = false;
                find.MatchSoundsLike = false;
                find.MatchAllWordForms = false;
                find.Execute();

                int i = 0;
                while (result.Find.Found && result.Text != value)
                {
                    result = range.Find(value);
                    if (i++ > 10)
                        throw new InvalidOperationException(string.Format("The value '{0}' exists in the range, but Word is unable to find it.", value));
                }

                return result;
            }
            finally
            {
                range.Document.ActiveWindow.View.ShowAll = ShowAll;
            }
        }

        public static Wd.Range Find(this Wd.Range range, Wd.WdFieldType type, string value = null)
        {
            var ShowAll = range.Document.ActiveWindow.View.ShowAll;
            var ShowFieldCodes = range.Document.ActiveWindow.View.ShowFieldCodes;

            try
            {
                range.Document.ActiveWindow.View.ShowAll = true;
                range.Document.ActiveWindow.View.ShowFieldCodes = true;
                value = string.Format("^d {0}^w{1}", type.AsGoToString(), value);

                Wd.Range result = range.Duplicate;
                Wd.Find find = result.Find;
                find.ClearFormatting();
                find.Text = value;
                find.Forward = true;
                find.Wrap = Wd.WdFindWrap.wdFindStop;
                find.Forward = true;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                find.MatchWildcards = false;
                find.MatchSoundsLike = false;
                find.MatchAllWordForms = false;
                find.Execute();

                return result;
            }
            finally
            {
                range.Document.ActiveWindow.View.ShowAll = ShowAll;
                range.Document.ActiveWindow.View.ShowFieldCodes = ShowFieldCodes;
            }
        }
    }
}
