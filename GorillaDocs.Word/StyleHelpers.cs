using GorillaDocs.libs.PostSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    [Log]
    public static class StyleHelpers
    {
        public static void SetStyle(this Wd.Range range, string styleName)
        {
            object style = range.Document.Styles[styleName];
            range.set_Style(ref style);
        }
        public static void SetStyle(this Wd.Selection selection, string styleName)
        {
            if (selection.Document.Styles.Exists(styleName))
            {
                object style = selection.Document.Styles[styleName];
                selection.set_Style(ref style);
            }
        }
        public static void SetStyle(this Wd.Paragraph para, string styleName)
        {
            object style = para.Range.Document.Styles[styleName];
            para.set_Style(ref style);
        }

        public static void SetStyle(this Wd.Range range, object styleType)
        {
            range.set_Style(ref styleType);
        }
        public static void SetStyle(this Wd.Selection selection, object styleType)
        {
            selection.set_Style(ref styleType);
        }
        public static void SetStyle(this Wd.Paragraph para, object styleType)
        {
            para.set_Style(ref styleType);
        }

        public static bool IsStyle(this Wd.Paragraph para, string styleName)
        {
            Wd.Style style = (Wd.Style)para.get_Style();
            return style.NameLocal == styleName;
        }
        public static bool IsStyle(this Wd.Range range, string styleName)
        {
            Wd.Style style = (Wd.Style)range.get_Style();
            return style.NameLocal == styleName;
        }
        public static bool IsStyle(this Wd.Selection selection, string styleName)
        {
            Wd.Style style = (Wd.Style)selection.Range.get_Style();
            return style.NameLocal == styleName;
        }

        [System.Diagnostics.DebuggerStepThrough]
        public static bool Exists(this Wd.Styles styles, string name)
        {
            try
            {
                Wd.Style style = styles[name];
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static void ImportStyle(this Wd.Template template, string style, string path)
        {
            try
            {
                template.Application.OrganizerCopy(path, template.FullName, style, Wd.WdOrganizerObject.wdOrganizerObjectStyles);
            }
            catch
            {
                Message.LogWarning("Unable to import style '{0}'", style);
            }
        }

        public static void ImportStyle(this Wd.Template template, Wd.WdBuiltinStyle style, string path)
        {
            try
            {
                template.Application.OrganizerCopy(path, template.FullName, template.Application.ActiveDocument.Styles[style].NameLocal, Wd.WdOrganizerObject.wdOrganizerObjectStyles);
            }
            catch
            {
                Message.LogWarning("Unable to import style '{0}'", style);
            }
        }

        public static void Hide(this Wd.Style style)
        {
            style.Visibility = true;
            style.UnhideWhenUsed = false;
        }
    }
}
