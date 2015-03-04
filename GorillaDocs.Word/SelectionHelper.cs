using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
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

        /// <summary>
        /// Prevents strange errors when selection is at the end of a document
        /// </summary>
        public static void InsertBreak_Safe(this Wd.Selection selection, Wd.WdBreakType BreakType = Wd.WdBreakType.wdSectionBreakNextPage)
        {
            var range = selection.Range;
            range.Text = "[Temp]";
            selection.InsertBreak(BreakType);
            range.MoveStart(Wd.WdUnits.wdCharacter, 1);
            range.Delete();
        }

        public static void DeleteSection(this Wd.Selection selection)
        {
            var doc = selection.Document;

            if (doc.Sections.Count == 1)
                throw new InvalidOperationException("This document only has one section. You can not delete it.");

            if (MessageBox.Show("Click 'OK' to remove the current section and its contents.\nOtherwise choose 'Cancel'", "Remove Section?", MessageBoxButton.OKCancel, MessageBoxImage.Question) == MessageBoxResult.OK)
            {
                selection.Sections[1].Delete();
                doc.UpdateAllFields();
            }
        }

    }
}
