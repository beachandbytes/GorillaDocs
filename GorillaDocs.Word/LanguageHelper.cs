using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class LanguageHelper
    {
        public static void SetProofingLanguage(this Wd.Document doc, CultureInfo culture)
        {
            var lcid = (Wd.WdLanguageID)culture.LCID;
            ((Wd.Style)doc.Styles[Wd.WdBuiltinStyle.wdStyleNormal]).LanguageID = lcid;
            doc.Range().LanguageID = lcid;
            doc.ContentControls.UpdateLanguageId(lcid);
        }
    }
}
