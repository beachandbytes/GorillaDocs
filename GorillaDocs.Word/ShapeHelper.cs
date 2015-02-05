using GorillaDocs.libs.PostSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    [Log]
    public static class ShapeHelper
    {
        public static string Title(this Wd.Shape shape, bool ThrowExceptionIfTitleNotDefined)
        {
            // Title didn't exist in Word 2003 compatible documents
            Wd.Document doc = (Wd.Document)shape.Parent;
            if (ThrowExceptionIfTitleNotDefined)
                return shape.Title;
            else
                return doc.CompatibilityMode <= (int)Wd.WdCompatibilityMode.wdWord2003 ? null : shape.Title;
        }

        public static IList<Wd.Shape> Shapes(this Wd.Document doc, Func<dynamic, bool> predicate = null)
        {
            var shapes = new List<Wd.Shape>();
            for (int i = 1; i <= doc.Shapes.Count; i++)
            {
                var shape = doc.Shapes[i];
                if (predicate == null || predicate(shape))
                    shapes.Add(shape);
            }
            return shapes;
        }

        public static IList<Wd.Shape> HeaderShapes(this Wd.Document doc, Func<dynamic, bool> predicate = null)
        {
            var shapes = new List<Wd.Shape>();
            for (int i = 1; i <= doc.Sections[1].Headers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count; i++)
            {
                var shape = doc.Sections[1].Headers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[i];
                if (predicate == null || predicate(shape))
                    shapes.Add(shape);
            }
            return shapes;
        }

        public static IList<Wd.Shape> FooterShapes(this Wd.Document doc, Func<dynamic, bool> predicate = null)
        {
            var shapes = new List<Wd.Shape>();
            for (int i = 1; i <= doc.Sections[1].Footers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count; i++)
            {
                var shape = doc.Sections[1].Footers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[i];
                if (predicate == null || predicate(shape))
                    shapes.Add(shape);
            }
            return shapes;
        }

        public static void Delete(this IList<Wd.Shape> shapes)
        {
            foreach (Wd.Shape shape in shapes)
                shape.Delete();
        }
    }
}
