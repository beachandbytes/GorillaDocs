using GorillaDocs.libs.PostSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    [Log]
    public static class RangeHelpers
    {
        public static void Remove(this Wd.Range range, string value)
        {
            while (range.Text.Contains(value))
            {
                var temp = range.Duplicate;
                var iPos = range.Text.IndexOf(value);
                temp.MoveStart(Wd.WdUnits.wdCharacter, iPos);
                temp.Collapse(Wd.WdCollapseDirection.wdCollapseStart);
                temp.MoveEnd(Wd.WdUnits.wdCharacter, value.Length);
                temp.Delete();
            }
        }

        public static void TypeText(this Wd.Range range, string value)
        {
            range.Text = value;
            range.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);
        }

        public static Wd.Table FindTable(this Wd.Range range, string description)
        {
            if (range.Document.CompatibilityMode < 14)
                throw new InvalidOperationException("This document must be in the Word 2010 DOCX file format. Create the document again if you want to use this functionality.");
            foreach (Wd.Table table in range.Tables)
                if (table.Descr == description)
                    return table;
            return null;
        }

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
        public static void CollapseAndExtendToEnd(this Wd.Range range)
        {
            range.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);
            range.End = range.Document.Range().End;
        }

        public static Wd.Range MoveToEnd(this Wd.Cell cell)
        {
            Wd.Range range = cell.Range;
            range.Collapse(Wd.WdCollapseDirection.wdCollapseStart);
            range.MoveUntil("\a");
            if (range.Characters.First.Previous().Text == "\r")
                range.Move(Wd.WdUnits.wdCharacter, -1);
            return range;
        }

        public static Wd.Range MoveToEndOfParagraph(this Wd.Range range)
        {
            range.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);
            range.MoveUntil("\r");
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
        public static Wd.Range Start(this Wd.Paragraph paragraph)
        {
            Wd.Range range = paragraph.Range;
            range.Collapse(Wd.WdCollapseDirection.wdCollapseStart);
            return range;
        }
        public static Wd.Range End(this Wd.Paragraph paragraph)
        {
            Wd.Range range = paragraph.Range;
            range.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);
            return range;
        }

        public static void RemoveFromEnd(this Wd.Range range, string value)
        {
            Wd.Range r = range.Duplicate;
            if (r.Text.EndsWith(value + "\r"))
            {
                r.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);
                r.Move(Wd.WdUnits.wdCharacter, -1);
                r.MoveStart(Wd.WdUnits.wdCharacter, -value.Length);
                r.Delete();
            }
        }

        public static void AddToEnd(this Wd.Range range, string value)
        {
            if (!range.Text.EndsWith(value + "\r"))
                range.Characters.Last.InsertBefore(value);
        }

        /// <summary>
        /// This method inserts the text provided and returns a collection of Wd.Range objects located at each occurance of {0} in the string.
        /// This method is useful when inserting fields in a string. It can be difficult to insert text, a field, and then text again. This is because the insertion point sometimes get 'stuck' inside the field.
        /// </summary>
        /// <param name="range">The insertion point.</param>
        /// <param name="input">A string similar to "My name is {0} and I am {0} years old."</param>
        /// <returns>A list of Wd.Range objects for each {0} in the input string.</returns>
        public static List<Wd.Range> InsertAndReturnCollection(this Wd.Range range, string input)
        {
            List<Wd.Range> ranges = new List<Wd.Range>();

            string[] delimiter = { "{0}" };
            string[] values = input.Split(delimiter, StringSplitOptions.None);

            range.Text = "";
            foreach (string value in values)
            {
                range.InsertAfter(value);
                Wd.Range tempRange = range.Duplicate;
                tempRange.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);
                ranges.Add(tempRange);
            }
            return ranges;
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

        public static void RestartNumbering(this Wd.Range range)
        {
            if (range.ListParagraphs.Count > 0)
            {
                var listformat = range.ListParagraphs[1].Range.ListFormat;
                listformat.ApplyListTemplateWithLevel(listformat.ListTemplate, false, Wd.WdListApplyTo.wdListApplyToWholeList, Wd.WdDefaultListBehavior.wdWord10ListBehavior);
            }
        }

        public static void DeleteLine(this Wd.Range range)
        {
            if (range.Text.EndsWith("\a"))
            {
                range.MoveEnd(Wd.WdUnits.wdCharacter, -1);
            }
            if (!range.Text.EndsWith("\a"))
                range.Delete();
        }

        public static Wd.Range MoveOutOfTable(this Wd.Range value, Wd.WdCollapseDirection collapse = Wd.WdCollapseDirection.wdCollapseStart)
        {
            Wd.Range range = value.Duplicate;
            if (collapse == Wd.WdCollapseDirection.wdCollapseStart)
            {
                if ((bool)range.Information[Wd.WdInformation.wdWithInTable])
                {
                    range = range.Tables[1].Range;
                    range.Collapse(Wd.WdCollapseDirection.wdCollapseStart);
                }
                if ((bool)range.Information[Wd.WdInformation.wdWithInTable])
                {
                    range.Select();
                    Wd.Selection selection = range.Application.Selection;
                    selection.InsertRowsAbove();
                    range = range.Tables[1].Split(2).Range;
                    range.MoveStart(Wd.WdUnits.wdCharacter,-2);
                    range.Collapse(Wd.WdCollapseDirection.wdCollapseStart);
                    range = range.Tables[1].Range;
                    range.Tables[1].Delete();
                }
                return range;
            }
            else
                throw new NotImplementedException();
        }

        public static void ScrollIntoView(this Wd.Range range, bool start = true)
        {
            range.Application.ActiveWindow.ScrollIntoView(range, start);
        }
    }
}
