﻿using GorillaDocs.libs.PostSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    [Log]
    public static class HeaderFooterHelper
    {
        public static IList<Wd.Shape> GetShapes(this Wd.HeaderFooter headerfooter, Func<dynamic, bool> predicate = null)
        {
            var shapes = new List<Wd.Shape>();
            for (int i = 1; i < headerfooter.Shapes.Count + 1; i++)
            {
                var shape = headerfooter.Shapes[i];
                if (shape.Anchor.InRange(headerfooter.Range))
                    if (predicate == null || predicate(shape))
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

        public static IList<Wd.HeaderFooter> Headers(this Wd.Sections sections, Func<dynamic, bool> predicate = null)
        {
            var headers = new List<Wd.HeaderFooter>();
            foreach (Wd.Section section in sections)
                foreach (Wd.HeaderFooter header in section.Headers)
                    if (predicate == null || predicate(header))
                        headers.Add(header);
            return headers;
        }
        public static IList<Wd.HeaderFooter> Footers(this Wd.Sections sections, Func<dynamic, bool> predicate = null)
        {
            var footers = new List<Wd.HeaderFooter>();
            foreach (Wd.Section section in sections)
                foreach (Wd.HeaderFooter footer in section.Footers)
                    if (predicate == null || predicate(footer))
                        footers.Add(footer);
            return footers;
        }

        public static IList<Wd.Field> Fields(this IList<Wd.HeaderFooter> headersFooters, Func<dynamic, bool> predicate = null)
        {
            var fields = new List<Wd.Field>();
            foreach (Wd.HeaderFooter headerFooter in headersFooters)
                foreach (Wd.Field field in headerFooter.Range.Fields)
                    if (predicate == null || predicate(field))
                        fields.Add(field);
            return fields;
        }

    }
}
