using GorillaDocs.libs.PostSharp;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    [Log]
    public static class TableOfContentsHelper
    {
        const string Bookmark_TOCRange = "TOCRange";
        const string Annexure_Style = "Annexure";
        const string Appendix_Style = "Appendix";
        const string Exhibit_Style = "Exhibit";
        const string Schedule_Style = "Schedule";

        public static Wd.TableOfContents Add(this Wd.TablesOfContents tocs, Wd.Range range)
        {
            range.Fields.Add(range, Wd.WdFieldType.wdFieldTOC, string.Format(@"\b ""{0}"" \o ""1-1"" \h \z ", Bookmark_TOCRange), false);
            Wd.TableOfContents toc = tocs[tocs.Count];
            toc.TabLeader = Wd.WdTabLeader.wdTabLeaderDots;
            return toc;
        }

        public static void Refresh(this Wd.TablesOfContents tocs, Wd.WdTabLeader tabLeader = Wd.WdTabLeader.wdTabLeaderDots, float Toc2LeftTab = 0)
        {
            UpdateTocBookmark(tocs);
            Wd.Document doc = tocs.Parent;
            doc.UpdateAllFields();
            tocs.UpdateListSeparators();
            foreach (Wd.TableOfContents toc in tocs)
            {
                toc.UpdateTabs(tabLeader, Toc2LeftTab);
                toc.TabLeader = tabLeader;
                toc.Update();
                if (doc.Styles[Wd.WdBuiltinStyle.wdStyleNormal].ParagraphFormat.ReadingOrder == Wd.WdReadingOrder.wdReadingOrderRtl)
                    toc.Range.ParagraphFormat.ReadingOrder = Wd.WdReadingOrder.wdReadingOrderRtl;
                Wd.Range range = toc.Range;
                range.Collapse(Wd.WdCollapseDirection.wdCollapseStart);
                range.Select();
            }
        }

        public static void UpdateListSeparators(this Wd.TablesOfContents tocs)
        {
            foreach (Wd.TableOfContents toc in tocs)
            {
                Wd.Range range = toc.Range;
                for (int i = 1; i <= range.Fields[1].Code.Characters.Count; i++)
                {
                    Wd.Range ch = range.Fields[1].Code.Characters[i];
                    if (ch.Text == "," && CultureInfo.CurrentCulture.TextInfo.ListSeparator == ";")
                        ch.Text = ";";
                    else if (ch.Text == ";" && CultureInfo.CurrentCulture.TextInfo.ListSeparator == ",")
                        ch.Text = ",";
                }
                toc.Update();
            }
        }

        public static void UpdateTabs(this Wd.TableOfContents toc, Wd.WdTabLeader tabLeader = Wd.WdTabLeader.wdTabLeaderDots, float Toc2LeftTab = 0)
        {
            Wd.Document doc = toc.Application.ActiveDocument;
            Wd.TabStops stops = null;
            float rightTab = 0;
            Wd.PageSetup pagesetup = toc.Range.Sections[1].PageSetup;
            if ((bool)toc.Range.Information[Wd.WdInformation.wdWithInTable])
                rightTab = toc.Range.Tables[1].Columns[1].Width - toc.Range.Tables[1].RightPadding - toc.Range.Tables[1].LeftPadding;
            else if (pagesetup.TextColumns.Count == 1)
                rightTab = pagesetup.PageWidth - pagesetup.RightMargin - pagesetup.LeftMargin;
            else
                rightTab = pagesetup.TextColumns.Width;

            if (doc.Styles.Exists("Contents Subheading"))
            {
                stops = doc.Styles["Contents Subheading"].ParagraphFormat.TabStops;
                stops.ClearAll();
                stops.Add(rightTab, Wd.WdTabAlignment.wdAlignTabRight, Wd.WdTabLeader.wdTabLeaderSpaces);
                stops.Add(rightTab + pagesetup.TextColumns.Spacing, Wd.WdTabAlignment.wdAlignTabRight, Wd.WdTabLeader.wdTabLeaderSpaces);
                stops.Add(pagesetup.PageWidth - pagesetup.RightMargin - pagesetup.LeftMargin, Wd.WdTabAlignment.wdAlignTabRight, Wd.WdTabLeader.wdTabLeaderSpaces);
            }

            stops = doc.Styles[Wd.WdBuiltinStyle.wdStyleTOC1].ParagraphFormat.TabStops;
            stops.ClearAll();
            stops.Add(rightTab, Wd.WdTabAlignment.wdAlignTabRight, tabLeader);

            stops = doc.Styles[Wd.WdBuiltinStyle.wdStyleTOC2].ParagraphFormat.TabStops;
            stops.ClearAll();
            if (Toc2LeftTab != 0)
                stops.Add(doc.Application.CentimetersToPoints(Toc2LeftTab), Wd.WdTabAlignment.wdAlignTabLeft); // Needed in DDR and Two column agreement
            stops.Add(rightTab, Wd.WdTabAlignment.wdAlignTabRight, tabLeader);

            stops = doc.Styles[Wd.WdBuiltinStyle.wdStyleTOC3].ParagraphFormat.TabStops;
            stops.ClearAll();
            stops.Add(rightTab, Wd.WdTabAlignment.wdAlignTabRight, tabLeader);

            stops = doc.Styles[Wd.WdBuiltinStyle.wdStyleTOC4].ParagraphFormat.TabStops;
            stops.ClearAll();
            stops.Add(rightTab, Wd.WdTabAlignment.wdAlignTabRight, tabLeader);
        }

        public static bool IsOneLevel(this Wd.TableOfContents toc)
        {
            Wd.Field field = toc.Range.Fields[1];
            if (field.Type == Wd.WdFieldType.wdFieldTOC)
                if (field.Code.Text.Contains(string.Format("Section Heading{0} 1", CultureInfo.CurrentCulture.TextInfo.ListSeparator))
                    || field.Code.Text.Contains("1-1"))
                    return true;
            return false;
        }
        public static bool IsTwoLevels(this Wd.TableOfContents toc)
        {
            Wd.Field field = toc.Range.Fields[1];
            if (field.Type == Wd.WdFieldType.wdFieldTOC)
                if (field.Code.Text.Contains(string.Format("Heading Style 1{0}2", CultureInfo.CurrentCulture.TextInfo.ListSeparator))
                    || field.Code.Text.Contains("1-2"))
                    return true;
            return false;
        }
        public static bool IsThreeLevels(this Wd.TableOfContents toc)
        {
            Wd.Field field = toc.Range.Fields[1];
            if (field.Type == Wd.WdFieldType.wdFieldTOC)
                if (field.Code.Text.Contains("1-3"))
                    return true;
            return false;
        }
        public static bool IsFourLevels(this Wd.TableOfContents toc)
        {
            Wd.Field field = toc.Range.Fields[1];
            if (field.Type == Wd.WdFieldType.wdFieldTOC)
                if (field.Code.Text.Contains("1-4"))
                    return true;
            return false;
        }

        public static void UpdateOneLevel(this Wd.TableOfContents toc)
        {
            toc.Range.Fields.Add(toc.Range, Wd.WdFieldType.wdFieldTOC, string.Format(@"\b ""{0}"" \o ""1-1"" \h \z ", Bookmark_TOCRange), false);
        }
        public static void UpdateTwoLevels(this Wd.TableOfContents toc)
        {
            toc.Range.Fields.Add(toc.Range, Wd.WdFieldType.wdFieldTOC, string.Format(@"\b ""{0}"" \o ""1-2"" \h \z ", Bookmark_TOCRange), false);
        }
        public static void UpdateThreeLevels(this Wd.TableOfContents toc)
        {
            toc.Range.Fields.Add(toc.Range, Wd.WdFieldType.wdFieldTOC, string.Format(@"\b ""{0}"" \o ""1-3"" \h \z ", Bookmark_TOCRange), false);
        }
        public static void UpdateFourLevels(this Wd.TableOfContents toc)
        {
            toc.Range.Fields.Add(toc.Range, Wd.WdFieldType.wdFieldTOC, string.Format(@"\b ""{0}"" \o ""1-4"" \h \z ", Bookmark_TOCRange), false);
        }
        public static void UpdateTocBookmark(Wd.TablesOfContents tocs)
        {
            var doc = (Wd.Document)tocs.Parent;
            if (tocs.Count > 0)
            {
                Wd.Range range = tocs.End();
                Wd.Range attachment = FindFirstAttachmentStyle(doc, range);
                if (attachment != null && attachment.Find.Found)
                    range.End = attachment.CollapseStart().Start - 2; // In case the Appendix starts with a content control
                else
                    range.End = doc.Range().End;
                doc.Bookmarks.Add(Bookmark_TOCRange, range);
            }
        }
        static Wd.Range FindFirstAttachmentStyle(Wd.Document doc, Wd.Range range)
        {
            Wd.Range result = null;
            result = FindFirstAttachmentStyleByName(doc, range, result, Annexure_Style);
            result = FindFirstAttachmentStyleByName(doc, range, result, Appendix_Style);
            result = FindFirstAttachmentStyleByName(doc, range, result, Exhibit_Style);
            result = FindFirstAttachmentStyleByName(doc, range, result, Schedule_Style);
            return result == null || !result.Find.Found ? null : result;
        }
        static Wd.Range FindFirstAttachmentStyleByName(Wd.Document doc, Wd.Range range, Wd.Range result, string style)
        {
            if (doc.Styles.Exists(style))
            {
                Wd.Range search = range.Duplicate;
                if (result == null || !result.Find.Found)
                    search.End = doc.Range().End;
                else
                    search.End = result.Start;

                Wd.Range result_temp = search.Find(null, doc.Styles[style]);

                if (result_temp.Find.Found && (result == null || result.Start > result_temp.Start))
                    return result_temp;
            }
            return result;
        }

        public static bool AttachmentExists(this Wd.TablesOfContents tocs)
        {
            foreach (Wd.TableOfContents toc in tocs)
            {
                Wd.Field field = toc.Range.Fields[1];
                if (field.Code.Text.Contains(Annexure_Style, StringComparison.OrdinalIgnoreCase) ||
                    field.Code.Text.Contains(Appendix_Style, StringComparison.OrdinalIgnoreCase) ||
                    field.Code.Text.Contains(Exhibit_Style, StringComparison.OrdinalIgnoreCase) ||
                    field.Code.Text.Contains(Schedule_Style, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

        public static void RemoveAttachment(this Wd.TablesOfContents tocs)
        {
            foreach (Wd.TableOfContents toc in tocs)
            {
                Wd.Field field = toc.Range.Fields[1];
                if (field.Code.Text.Contains(Annexure_Style, StringComparison.OrdinalIgnoreCase) ||
                    field.Code.Text.Contains(Appendix_Style, StringComparison.OrdinalIgnoreCase) ||
                    field.Code.Text.Contains(Exhibit_Style, StringComparison.OrdinalIgnoreCase) ||
                    field.Code.Text.Contains(Schedule_Style, StringComparison.OrdinalIgnoreCase))
                    toc.Delete();
            }
        }

        public static void AddAttachment(this Wd.TablesOfContents tocs)
        {
            if (tocs.Count > 0)
            {
                Wd.Range range = tocs.End();
                range.Fields.Add(range, Wd.WdFieldType.wdFieldTOC, string.Format(@"\h \z \t ""{0},1,{1},1,{2},1,{3},1""", Annexure_Style, Appendix_Style, Exhibit_Style, Schedule_Style), false);
            }
        }

        public static IList<Wd.TableOfContents> TablesOfContents(this Wd.Document doc)
        {
            var tocs = new List<Wd.TableOfContents>();
            foreach (Wd.TableOfContents toc in doc.TablesOfContents)
                tocs.Add(toc);
            return tocs;
        }

        public static void Update(this IList<Wd.TableOfContents> tocs, bool UpdateBookmarks = false)
        {
            if (UpdateBookmarks && tocs.Count > 0)
                UpdateTocBookmark(tocs[tocs.Count - 1]);
            foreach (Wd.TableOfContents toc in tocs)
                if (UpdateBookmarks && !toc.ContainsTocBookmark())
                    toc.UseTocBookmark();
                else
                    toc.Update();
        }

        public static bool ContainsTocBookmark(this Wd.TableOfContents toc)
        {
            var field = toc.Range.Fields[1];
            if (field.Type == Wd.WdFieldType.wdFieldTOC)
                if (field.Code.Text.Contains(string.Format(@"\b ""{0}""", Bookmark_TOCRange)))
                    return true;
            return false;
        }

        public static void UseTocBookmark(this Wd.TableOfContents toc)
        {
            var field = toc.Range.Fields[1];
            if (field.Type == Wd.WdFieldType.wdFieldTOC)
                toc.Range.Fields.Add(toc.Range, Wd.WdFieldType.wdFieldEmpty, string.Format(@"{0} \b ""{1}""", field.Code.Text, Bookmark_TOCRange), false);
        }

        public static void UpdateTocBookmark(this Wd.TableOfContents toc)
        {
            Wd.Range range = toc.Range.CollapseEnd();
            Wd.Document doc = range.Document;
            Wd.Range attachment = FindFirstAttachmentStyle(doc, range);
            if (attachment != null && attachment.Find.Found)
                range.End = attachment.CollapseStart().Start - 2; // In case the Appendix starts with a content control
            else
                range.End = doc.Range().End;
            doc.Bookmarks.Add(Bookmark_TOCRange, range);
        }
    }
}
