using System;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class TocHelper
    {
        public static void RefreshToc(this BaseDocument doc) { doc.TablesOfContents.Refresh(); }

        public static void ToggleTOCLevel(this BaseDocument doc)
        {
            for (int i = doc.TablesOfContents.Count; i > 0; i--)
            {
                Wd.TableOfContents toc = doc.TablesOfContents[i];
                if (toc.IsFourLevels())
                    toc.UpdateOneLevel();
                else if (toc.IsThreeLevels())
                    toc.UpdateFourLevels();
                else if (toc.IsTwoLevels())
                    toc.UpdateThreeLevels();
                else if (toc.IsOneLevel())
                    toc.UpdateTwoLevels();
            }
            doc.RefreshToc();
        }

        public static void ToggleTOCAppendices(this BaseDocument doc)
        {
            if (doc.TablesOfContents.AttachmentExists())
                doc.TablesOfContents.RemoveAttachment();
            else
                doc.TablesOfContents.AddAttachment();
        }
    }
}
