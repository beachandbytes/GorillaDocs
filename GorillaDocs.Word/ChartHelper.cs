using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class ChartHelper
    {
        public static Wd.Chart GetSelectedChartOrAddNew(this Wd.Selection selection)
        {
            if (selection.Type == Wd.WdSelectionType.wdSelectionInlineShape && selection.InlineShapes[1].Type == Wd.WdInlineShapeType.wdInlineShapeChart)
                return selection.InlineShapes[1].Chart;
            else
                return selection.InlineShapes.AddChart().Chart;
        }
    }
}
