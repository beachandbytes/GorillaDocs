using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class PaperSizeHelper
    {
        public static void UpdatePaperSize(this Wd.Document doc, Wd.WdPaperSize paperSize) { doc.Range().UpdatePaperSize(paperSize); }
        public static void UpdatePaperSize(this Wd.Range range, Wd.WdPaperSize paperSize)
        {
            try
            {
                foreach (Wd.Section section in range.Sections)
                    section.PageSetup.SetPaperSize(paperSize);
            }
            finally
            {
                range.Application.ScreenRefresh();
            }
        }
    }
}
