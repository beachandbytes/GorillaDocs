using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;
using O = Microsoft.Office.Core;

namespace GorillaDocs.Word
{
    public static class ToggleHeaderFooterHelper
    {
        const int Visible = 0;
        const int Hidden = -1;
        public static void ToggleHeaderFooter(this Wd.Document doc)
        {
            if (doc.Styles[Wd.WdBuiltinStyle.wdStyleHeader].Font.Hidden == Visible)
            {
                doc.Styles[Wd.WdBuiltinStyle.wdStyleHeader].Font.Hidden = Hidden;
                doc.Styles[Wd.WdBuiltinStyle.wdStyleFooter].Font.Hidden = Hidden;
                doc.HeaderShapes().SetVisible(O.MsoTriState.msoFalse);
            }
            else
            {
                doc.Styles[Wd.WdBuiltinStyle.wdStyleHeader].Font.Hidden = Visible;
                doc.Styles[Wd.WdBuiltinStyle.wdStyleFooter].Font.Hidden = Visible;
                doc.HeaderShapes().SetVisible(O.MsoTriState.msoTrue);
            }
        }

        public static void SetVisible(this IList<Wd.Shape> shapes, O.MsoTriState Visible)
        {
            foreach (Wd.Shape shape in shapes)
                shape.Visible = Visible;
        }
    }
}