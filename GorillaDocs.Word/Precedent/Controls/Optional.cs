using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.Precedent.Controls
{
    public class Optional : PrecedentControl
    {
        public Optional(Wd.ContentControl control) : base(control) { }

        public static bool Test(Wd.ContentControl control)
        {
            return control.Type == Wd.WdContentControlType.wdContentControlRichText &&
                control.PrecedentInstruction() != null &&
                control.PrecedentInstruction().Command == "Optional";
        }

        public override void Process()
        {
            var result = PrecedentExpression1.Resolve(control.PrecedentInstruction().Expression, control.PrecedentInstruction().ExpressionData(control.Range.Document));
            if (result == true)
                try
                {
                    DeleteInstructions(control);
                    //control.Range.ContentControls.ConvertMappedWithValueToText();
                    control.Delete();
                }
                catch (Exception ex)
                {
                    Message.LogError(ex);
                }
            else
                DeleteParagraphIfEmpty1(control);
        }

        static void DeleteInstructions(Wd.ContentControl control)
        {
            try
            {
                foreach (Wd.ContentControl item in control.Range.ContentControls)
                    if (!string.IsNullOrEmpty(item.Title) && item.Title.ToLower() == "instruction")
                        item.DeleteParagraphIfEmpty();
            }
            catch (Exception ex)
            {
                Message.LogError(ex);
            }
        }

        static void DeleteParagraphIfEmpty1(Wd.ContentControl control)
        {
            Wd.Range range = control.Range;
            control.Delete(true);
            if (range.Paragraphs[1].IsEmpty())
                range.Paragraphs[1].Range.Delete();
        }
    }
}
