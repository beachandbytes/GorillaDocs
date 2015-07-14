using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class ParagraphHelper
    {
        public static Wd.Paragraph InsertBefore(this Wd.Paragraph para)
        {
            var range = para.Range.CollapseStart();
            range.InsertParagraphBefore();
            return range.CollapseStart().Paragraphs[1];
        }

        public static Wd.Range InsertParagraph(this Wd.Range range, string text, string style = null)
        {
            if (!range.Paragraphs[1].IsEmpty())
                range = range.Paragraphs[1].InsertBefore().Range.CollapseStart();
            range.Text = text;
            if (!string.IsNullOrEmpty(style))
                range.SetStyle(style);
            return range;
        }

        public static Wd.Range InsertParagraphAfter(this Wd.Range range, Wd.WdCollapseDirection collapse)
        {
            range.InsertParagraphAfter();
            if (collapse == Wd.WdCollapseDirection.wdCollapseStart)
                return range.Paragraphs.Last.Next().Range.CollapseStart();
            else
                return range.Paragraphs.Last.Next().Range.CollapseEnd();
        }

        public static Wd.Range ExpandParagraph(this Wd.Range range)
        {
            range.Expand(Wd.WdUnits.wdParagraph);
            return range;
        }

    }
}
