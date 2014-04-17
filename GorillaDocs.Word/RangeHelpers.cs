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

    }
}
