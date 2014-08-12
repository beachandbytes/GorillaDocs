using GorillaDocs.libs.PostSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    [Log]
    public static class SectionHelper
    {
        public static Wd.Section AddSection(this Wd.Range range, Wd.WdBreakType BreakType = Wd.WdBreakType.wdSectionBreakNextPage)
        {
            range.InsertBreak(BreakType);
            Wd.Section section = range.Sections[1];
            foreach (Wd.HeaderFooter header in section.Headers)
                header.LinkToPrevious = false;
            foreach (Wd.HeaderFooter footer in section.Footers)
                footer.LinkToPrevious = false;
            return section.Previous();
        }
        public static Wd.Section AddSection(this Wd.Selection selection, Wd.WdBreakType BreakType = Wd.WdBreakType.wdSectionBreakNextPage)
        {
            return selection.Range.AddSection(BreakType);
        }

        public static bool IsStartOfSection(this Wd.Selection selection)
        {
            if (selection.Characters.First.Start == selection.Document.Characters.First.Start)
                return true;
            if (selection.Characters.First.Sections[1].Index != selection.Characters.First.Previous().Sections[1].Index)
                return true;
            return false;
        }

        public static void ApplyPreviousPageSetup(this Wd.Section section)
        {
            //TODO: Add more settings as they are needed
            if (section.Index != 1)
            {
                section.PageSetup.TopMargin = section.Previous().PageSetup.TopMargin;
                section.PageSetup.LeftMargin = section.Previous().PageSetup.LeftMargin;
                section.PageSetup.RightMargin = section.Previous().PageSetup.RightMargin;
                section.PageSetup.BottomMargin = section.Previous().PageSetup.BottomMargin;
                section.PageSetup.VerticalAlignment = section.Previous().PageSetup.VerticalAlignment;
            }
        }

        public static void Delete(this Wd.Section section)
        {
            var doc = section.Range.Document;
            if (doc.Sections.Count == 1)
                throw new InvalidOperationException("This document only has one section. You can not delete it.");

            section.ApplyPreviousPageSetup();

            if (section.Index == doc.Sections.Count)
            {
                foreach (Wd.HeaderFooter hf in section.Headers)
                    try
                    {
                        hf.LinkToPrevious = true;
                        hf.LinkToPrevious = false;
                    }
                    catch (Exception ex) { Message.LogError(ex); }
                foreach (Wd.HeaderFooter hf in section.Footers)
                    try
                    {
                        hf.LinkToPrevious = true;
                        hf.LinkToPrevious = false;
                    }
                    catch (Exception ex) { Message.LogError(ex); }
                Wd.Range rng = section.Range;
                rng.MoveStart(Wd.WdUnits.wdCharacter, -1);
                rng.Delete();
            }
            else
            {
                section.RestartNumbering();
                section.Range.Delete();
            }
        }
        public static void DeleteSection(this Wd.Document doc, int i)
        {
            doc.Sections[i].Delete();
        }

        public static void ToggleOrientation(this Wd.Section section)
        {
            Wd.PageSetup pageSetup = section.PageSetup;
            float topMargin = pageSetup.TopMargin;
            float leftMargin = pageSetup.LeftMargin;
            float rightMargin = pageSetup.RightMargin;
            float bottomMargin = pageSetup.BottomMargin;

            if (pageSetup.Orientation == Wd.WdOrientation.wdOrientPortrait)
                pageSetup.Orientation = Wd.WdOrientation.wdOrientLandscape;
            else
                pageSetup.Orientation = Wd.WdOrientation.wdOrientPortrait;

            pageSetup.TopMargin = topMargin;
            pageSetup.LeftMargin = leftMargin;
            pageSetup.RightMargin = rightMargin;
            pageSetup.BottomMargin = bottomMargin;
        }

        public static void ToggleOrientation(this Wd.Section section, float wideMargin, float narrowMargin)
        {
            section.ToggleOrientation();
            Wd.PageSetup pageSetup = section.PageSetup;
            if (pageSetup.RightMargin == wideMargin)
                section.ReJigRightMarginObjects(wideMargin - narrowMargin);
            else
                section.ReJigRightMarginObjects();
        }

        public static void ReJigRightMarginObjects(this Wd.Section section, float narrowMargin, float wideMargin)
        {
            if (section.PageSetup.RightMargin == narrowMargin)
                section.ReJigRightMarginObjects();
            else
                section.ReJigRightMarginObjects(wideMargin - narrowMargin);
        }
        public static void ReJigRightMarginObjects(this Wd.Section section, float offset = 0)
        {
            foreach (Wd.HeaderFooter header in section.Headers)
                if (header.Exists)
                    ReJigRightMarginObjects_HeaderFooter(header, section.PageSetup.EditableWidth() + offset, offset);
            foreach (Wd.HeaderFooter footer in section.Footers)
                if (footer.Exists)
                    ReJigRightMarginObjects_HeaderFooter(footer, section.PageSetup.EditableWidth() + offset, offset);
        }
        static void ReJigRightMarginObjects_HeaderFooter(Wd.HeaderFooter item, float tableWidth, float offset)
        {
            foreach (Wd.Shape shape in item.GetShapes())
                if (shape.RelativeHorizontalPosition == Wd.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionRightMarginArea)
                    shape.Left = -shape.Width + offset;
            foreach (Wd.Table table in item.Range.Tables)
                if (offset == 0)
                    table.AutoFitBehavior(Wd.WdAutoFitBehavior.wdAutoFitWindow);
                else
                {
                    table.PreferredWidthType = Wd.WdPreferredWidthType.wdPreferredWidthPoints;
                    table.PreferredWidth = tableWidth;
                }
        }

        public static void InsertLogoInHeaders(this Wd.Section section, string path, string name, float top)
        {
            foreach (Wd.HeaderFooter header in section.Headers)
            {
                Wd.Shape shape = header.Shapes.AddPicture(path);
                shape.RelativeVerticalPosition = Wd.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                shape.RelativeHorizontalPosition = Wd.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionRightMarginArea;
                shape.Name = name;
                shape.Top = top;
                shape.Left = -shape.Width;
            }
        }

        public static void RestartNumbering(this Wd.Section section)
        {
            var footer = section.Footers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary];
            footer.PageNumbers.RestartNumberingAtSection = true;
            footer.PageNumbers.StartingNumber = 1;
        }
        public static void ContinueNumbering(this Wd.Section section)
        {
            if (section.Index != 1)
            {
                Wd.HeaderFooter hf = section.Footers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                hf.PageNumbers.RestartNumberingAtSection = false;
                hf.PageNumbers.StartingNumber = 0;
            }
        }

        public static bool IsArabicNumbering(this Wd.Section section)
        {
            return section.Footers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.NumberStyle == Wd.WdPageNumberStyle.wdPageNumberStyleArabic;
        }
        public static bool IsRomanNumbering(this Wd.Section section)
        {
            return section.Footers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.NumberStyle == Wd.WdPageNumberStyle.wdPageNumberStyleLowercaseRoman ||
                section.Footers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.NumberStyle == Wd.WdPageNumberStyle.wdPageNumberStyleUppercaseRoman;
        }

        public static Wd.Section Previous(this Wd.Section section)
        {
            if (section.Index == 1)
                return null;
            else
                return section.Range.Document.Sections[section.Index - 1];
        }
        public static Wd.Section Next(this Wd.Section section)
        {
            if (section.Index + 1 > section.Range.Document.Sections.Count)
                return null;
            else
                return section.Range.Document.Sections[section.Index + 1];
        }

        public static void ClearHeaders(this Wd.Section section)
        {
            Wd.Section nextSection = section.Next();
            foreach (Wd.HeaderFooter header in section.Headers)
            {
                if (nextSection != null && nextSection.Headers[header.Index].Exists)
                    nextSection.Headers[header.Index].LinkToPrevious = false;
                header.LinkToPrevious = false;
                header.Range.Delete();
            }
        }
        public static void ClearFooters(this Wd.Section section)
        {
            Wd.Section nextSection = section.Next();
            foreach (Wd.HeaderFooter footer in section.Footers)
            {
                if (nextSection != null && nextSection.Footers[footer.Index].Exists)
                    nextSection.Footers[footer.Index].LinkToPrevious = false;
                footer.LinkToPrevious = false;
                footer.Range.Delete();
            }
        }

        public static bool AreHeaderFootersEmpty(this Wd.Sections sections)
        {
            foreach (Wd.Section section in sections)
                if (!section.AreHeaderFootersEmpty())
                    return false;
            return true;
        }
        public static bool AreHeaderFootersEmpty(this Wd.Section section)
        {
            foreach (Wd.HeaderFooter header in section.Headers)
                if (header.Range.Characters.Count != 1)
                    return false;
            foreach (Wd.HeaderFooter footer in section.Footers)
                if (footer.Range.Characters.Count != 1)
                    return false;
            return true;
        }

        public static bool AreHeadersEmpty(this Wd.Sections sections)
        {
            foreach (Wd.Section section in sections)
                if (!section.AreHeadersEmpty())
                    return false;
            return true;
        }
        public static bool AreHeadersEmpty(this Wd.Section section)
        {
            foreach (Wd.HeaderFooter header in section.Headers)
                if (header.Range.Characters.Count != 1 || header.ContainsShapes())
                    return false;
            return true;
        }

        public static bool AreFootersEmpty(this Wd.Sections sections)
        {
            foreach (Wd.Section section in sections)
                if (!section.AreFootersEmpty())
                    return false;
            return true;
        }
        public static bool AreFootersEmpty(this Wd.Section section)
        {
            foreach (Wd.HeaderFooter footer in section.Footers)
                if (footer.Range.Characters.Count != 1)
                    return false;
            return true;
        }

        public static bool ContainsColumns(this Wd.Section section)
        {
            return section.PageSetup.TextColumns.Count > 1;
        }

        public static Wd.PageNumbers PageNumbers(this Wd.Section section)
        {
            return section.Footers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers;
        }

        public static bool ContainsShapeNamed(this Wd.Section section, string name)
        {
            var doc = (Wd.Document)section.Parent;
            if (section.Range.ContainsShapeNamed(doc.Shapes, name))
                return true;

            foreach (Wd.HeaderFooter header in section.Headers)
                if (header.Range.ContainsShapeNamed(header.Shapes, name))
                    return true;

            foreach (Wd.HeaderFooter footer in section.Footers)
                if (footer.Range.ContainsShapeNamed(footer.Shapes, name))
                    return true;

            return false;
        }
        static bool ContainsShapeNamed(this Wd.Range range, Wd.Shapes shapes, string name)
        {
            for (int i = 1; i < shapes.Count + 1; i++)
            {
                var shape = shapes[i];
                if (shape.Anchor.InRange(range) && shape.Name == name)
                    return true;
            }
            return false;
        }

        public static bool ContainsShapeWithTitle(this Wd.Section section, string title)
        {
            var doc = (Wd.Document)section.Parent;
            if (section.Range.ContainsShapeWithTitle(doc.Shapes, title))
                return true;

            foreach (Wd.HeaderFooter header in section.Headers)
                if (header.Range.ContainsShapeWithTitle(header.Shapes, title))
                    return true;

            foreach (Wd.HeaderFooter footer in section.Footers)
                if (footer.Range.ContainsShapeWithTitle(footer.Shapes, title))
                    return true;

            return false;
        }
        static bool ContainsShapeWithTitle(this Wd.Range range, Wd.Shapes shapes, string title)
        {
            for (int i = 1; i < shapes.Count + 1; i++)
            {
                var shape = shapes[i];
                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox &&
                    shape.Anchor.InRange(range) && shape.Title == title)
                    return true;
            }
            return false;
        }

        public static void WrapInNewSection(this Wd.Range range, Wd.WdBreakType breakType = Wd.WdBreakType.wdSectionBreakContinuous)
        {
            if (!range.IsStartOfSection())
            {
                range.Document.Range(range.Start, range.Start).InsertBreak(breakType);
                range.Start++;
            }
            if (!range.IsEndOfSection())
                range.Document.Range(range.End, range.End).InsertBreak(breakType);
        }

        public static bool IsStartOfSection(this Wd.Range range)
        {
            if (range.Start == 0)
                return true;
            else
                return range.Document.Range(range.Start - 1, range.Start).Characters[1].Text == ((char)12).ToString();
        }
        public static bool IsEndOfSection(this Wd.Range range)
        {
            return range.IsEndOfDocument() ||
                (range.Document.Range(range.End, range.End + 1).Characters[1].Text == ((char)12).ToString());
        }

        public static bool IsEndOfDocument(this Wd.Range range)
        {
            return range.End + 1 == range.Document.Range().End;
        }
    }
}
