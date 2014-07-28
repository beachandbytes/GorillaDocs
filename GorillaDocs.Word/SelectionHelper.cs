using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class SelectionHelper
    {
        public static void ExtendToEndOfStyle(this Wd.Selection selection)
        {
            if (selection.Start == selection.End)
                selection.Characters.Last.Select();
            string name = ((Wd.Style)selection.get_Style()).NameLocal;

            Wd.Range ch = selection.Characters.Last.Next();
            if (ch == null)
                return;
            while (ch.IsStyle(name))
            {
                selection.MoveEnd(Wd.WdUnits.wdCharacter, 1);
                ch = selection.Characters.Last.Next();
                if (ch == null)
                    return;
            }
        }

        public static void ExtendToStartOfStyle(this Wd.Selection selection)
        {
            Wd.Range ch = selection.Characters.First.Previous();
            string name = ((Wd.Style)selection.get_Style()).NameLocal;

            if (ch == null)
                return;
            while (ch.IsStyle(name))
            {
                selection.MoveStart(Wd.WdUnits.wdCharacter, -1);
                ch = selection.Characters.First.Previous();
                if (ch == null)
                    return;
            }
        }

        public static void RemoveQuotes(this Wd.Selection selection)
        {
            Wd.Range range = selection.Range;
            Wd.Range f = range.Characters.First;
            Wd.Range l = range.Characters.Last;
            if (f.Text == "\"")
                f.Delete();
            if (l.Text == "\"")
                l.Delete();
            range.Select();
        }

    }
}
