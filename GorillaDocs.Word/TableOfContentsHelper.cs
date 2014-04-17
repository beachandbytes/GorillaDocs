using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class TableOfContentsHelper
    {
        public static void Refresh(this Wd.TablesOfContents tocs, Wd.WdTabLeader tabLeader = Wd.WdTabLeader.wdTabLeaderDots, float Toc2LeftTab = 0)
        {
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

        public static void UpdateOneLevel(this Wd.TableOfContents toc)
        {
            toc.Range.Fields.Add(toc.Range, Wd.WdFieldType.wdFieldTOC, @"\o ""1-1"" \h \z ", false);
        }
        public static void UpdateTwoLevels(this Wd.TableOfContents toc)
        {
            toc.Range.Fields.Add(toc.Range, Wd.WdFieldType.wdFieldTOC, @"\o ""1-2"" \h \z ", false);
        }
        public static void UpdateThreeLevels(this Wd.TableOfContents toc)
        {
            toc.Range.Fields.Add(toc.Range, Wd.WdFieldType.wdFieldTOC, @"\o ""1-3"" \h \z ", false);
        }
    }
}
