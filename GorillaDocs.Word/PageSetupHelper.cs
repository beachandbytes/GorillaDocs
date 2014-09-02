using GorillaDocs.libs.PostSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    [Log]
    public static class PageSetupHelper
    {
        public static float EditableWidth(this Wd.PageSetup pageSetup)
        {
            return pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin;
        }

        public static void SetPaperSize(this Wd.PageSetup pageSetup, Wd.WdPaperSize paperSize)
        {
            Wd.WdOrientation orientation = pageSetup.Orientation;
            if (pageSetup.PaperSize != paperSize)
            {
                pageSetup.PaperSize = paperSize;
                if (pageSetup.Orientation != orientation)
                {
                    pageSetup.Orientation = orientation;
                    var temp = pageSetup.TopMargin;
                    pageSetup.TopMargin = pageSetup.LeftMargin;
                    pageSetup.LeftMargin = pageSetup.BottomMargin;
                    pageSetup.BottomMargin = pageSetup.RightMargin;
                    pageSetup.RightMargin = temp;
                }
            }
        }
    }
}
