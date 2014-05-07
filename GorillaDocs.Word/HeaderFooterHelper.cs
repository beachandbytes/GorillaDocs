using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class HeaderFooterHelper
    {
        public static List<Wd.Shape> GetShapes(this Wd.HeaderFooter headerfooter)
        {
            var shapes = new List<Wd.Shape>();
            for (int i = 1; i < headerfooter.Shapes.Count + 1; i++)
            {
                var shape = headerfooter.Shapes[i];
                if (shape.Anchor.InRange(headerfooter.Range))
                    shapes.Add(shape);
            }
            return shapes;
        }

        public static bool ContainsShapes(this Wd.HeaderFooter headerfooter)
        {
            for (int i = 1; i < headerfooter.Shapes.Count + 1; i++)
            {
                var shape = headerfooter.Shapes[i];
                if (shape.Anchor.InRange(headerfooter.Range))
                    return true;
            }
            return false;
        }
    }
}
