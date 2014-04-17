using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class SectionHelper
    {
        public static void Delete(this Wd.Section section)
        {
            var doc = section.Range.Document;
            if (doc.Sections.Count == 1)
                throw new InvalidOperationException("This document only has one section. You can not delete it.");

            if (section.Index == doc.Sections.Count)
            {
                // TODO: Bookmarks in the second last header are deleted when the last section takes over. Need some way of copying them between?
                foreach (Wd.HeaderFooter hf in section.Headers)
                {
                    hf.LinkToPrevious = true;
                    hf.LinkToPrevious = false;
                }
                foreach (Wd.HeaderFooter hf in section.Footers)
                {
                    hf.LinkToPrevious = true;
                    hf.LinkToPrevious = false;
                }
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

        public static void RestartNumbering(this Wd.Section section)
        {
            var doc = section.Range.Document;
            if (section.Index == doc.Sections.Count)
                throw new InvalidOperationException("Can not set RestartNumberingAtSection for the last section");

            if (section.Index != 1)
            {
                var footer = section.Footers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                var footer_Next = doc.Sections[section.Index + 1].Footers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                if (footer.PageNumbers.RestartNumberingAtSection && !footer_Next.PageNumbers.RestartNumberingAtSection)
                {
                    footer_Next.PageNumbers.RestartNumberingAtSection = true;
                    footer_Next.PageNumbers.StartingNumber = 1;
                }
            }
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

        public static void ContinueNumbering(this Wd.Section section)
        {
            Wd.Document doc = section.Application.ActiveDocument;
            if (section.Index == doc.Sections.Count)
                throw new InvalidOperationException("Can not set ContinueNumberingAtSection for the last section");

            if (section.Index != 1)
            {
                Wd.HeaderFooter hf = section.Footers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                hf.PageNumbers.RestartNumberingAtSection = false;
                hf.PageNumbers.StartingNumber = 0;

                Wd.HeaderFooter hf_Next = doc.Sections[section.Index + 1].Footers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                hf_Next.PageNumbers.RestartNumberingAtSection = false;
                hf_Next.PageNumbers.StartingNumber = 0;
            }
        }

        public static bool IsArabicNumbering(this Wd.Section section)
        {
            return section.Footers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.NumberStyle == Wd.WdPageNumberStyle.wdPageNumberStyleArabic;
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
            for (int i = 1; i < doc.Shapes.Count + 1; i++)
            {
                var shape = doc.Shapes[i];
                if (shape.Anchor.InRange(section.Range) && shape.Name == name)
                    return true;
            }
            return false;
        }

        public static void WrapInNewSection(this Wd.Range range)
        {
            if (!range.IsStartOfSection())
            {
                range.Document.Range(range.Start, range.Start).InsertBreak(Wd.WdBreakType.wdSectionBreakContinuous);
                range.Start++;
            }
            if (!range.IsEndOfSection())
                range.Document.Range(range.End, range.End).InsertBreak(Wd.WdBreakType.wdSectionBreakContinuous);
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
            return range.Document.Range(range.End, range.End + 1).Characters[1].Text == ((char)12).ToString();
        }
    }
}
