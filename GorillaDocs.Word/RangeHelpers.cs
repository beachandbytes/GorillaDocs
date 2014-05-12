using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class RangeHelpers
    {
        public static Wd.Range CollapseStart(this Wd.Range range)
        {
            range.Collapse(Wd.WdCollapseDirection.wdCollapseStart);
            return range;
        }
        public static Wd.Range CollapseEnd(this Wd.Range range)
        {
            range.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);
            return range;
        }

        public static Wd.Range Start(this Wd.Document doc)
        {
            Wd.Range range = doc.Range();
            range.Collapse(Wd.WdCollapseDirection.wdCollapseStart);
            return range;
        }
        public static Wd.Range End(this Wd.Document doc)
        {
            Wd.Range range = doc.Range();
            range.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);
            return range;
        }
        public static Wd.Range Start(this Wd.Section section)
        {
            Wd.Range range = section.Range;
            range.Collapse(Wd.WdCollapseDirection.wdCollapseStart);
            return range;
        }
        public static Wd.Range End(this Wd.Section section)
        {
            Wd.Range range = section.Range;
            range.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);
            return range;
        }
        public static Wd.Range Start(this Wd.TablesOfContents tocs)
        {
            Wd.Range range = tocs[1].Range;
            range.Collapse(Wd.WdCollapseDirection.wdCollapseStart);
            return range;
        }
        public static Wd.Range End(this Wd.TablesOfContents tocs)
        {
            Wd.Range range = tocs[tocs.Count].Range;
            range.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);
            return range;
        }

        public static Wd.Range Find(this Wd.Range searchRange, string value = null, Wd.Style style = null, bool wildcards = false)
        {
            bool showall = searchRange.Application.ActiveWindow.ActivePane.View.ShowAll;
            searchRange.Application.ActiveWindow.ActivePane.View.ShowAll = true;

            try
            {
                var range = searchRange.Duplicate;
                range.Find.ClearFormatting();
                if (style != null)
                    range.Find.set_Style(style);
                if (value != null)
                    range.Find.Text = value;
                range.Find.Forward = true;
                range.Find.Wrap = Wd.WdFindWrap.wdFindStop;
                range.Find.Format = true;
                range.Find.MatchCase = false;
                range.Find.MatchWholeWord = false;
                range.Find.MatchWildcards = wildcards;
                range.Find.MatchSoundsLike = false;
                range.Find.MatchAllWordForms = false;
                range.Find.Execute();
                return range;
            }
            finally
            {
                searchRange.Application.ActiveWindow.ActivePane.View.ShowAll = showall;
            }
        }

        public static bool IsCollapsed(this Wd.Range range)
        {
            return range.Start == range.End;
        }

        public static void DeleteParagraphFromRange(this Wd.Range range)
        {
            Wd.Range prev = range.Duplicate;
            try
            {
                range.Expand(Wd.WdUnits.wdParagraph);
                if (!range.ContainsTableCell())
                    range.Delete();
            }
            catch
            {
                prev.Delete();
            }
        }

    }
}
