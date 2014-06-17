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
    }
}
