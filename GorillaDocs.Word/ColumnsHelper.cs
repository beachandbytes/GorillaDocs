using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class ColumnsHelper
    {
        public static void ToggleColumns(this Wd.Section section)
        {
            section.Range.ToggleColumns();
        }
        public static void ToggleColumns(this Wd.Range range)
        {
            if (range.PageSetup.TextColumns.Count == 1)
                range.PageSetup.TextColumns.SetCount(2);
            else
                range.PageSetup.TextColumns.SetCount(1);
        }

        public static void ToggleColumns(this Wd.Selection selection)
        {
            if (selection.Characters.Count == 1)
                selection.Range.ToggleColumns();
            else
            {
                Wd.Range range = selection.Range;
                range.WrapInNewSection();
                range.ToggleColumns();
            }
        }

    }
}
