using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class CaptionHelper
    {
        public static void AddCaption(this Wd.Range range, string Label = "Figure", string PlaceholderText = "Enter caption here.")
        {
            range.InsertCaption(Label);
            range.MoveToEndOfParagraph().TypeText(": ");
            var contentControl = range.MoveToEndOfParagraph().ContentControls.Add_Safely();
            contentControl.SetPlaceholderText(null, null, PlaceholderText);
            contentControl.Temporary = true;
        }
    }
}
