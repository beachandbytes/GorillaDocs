using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.Precedent.Controls
{
    class DeleteControlIf : PrecedentControl
    {
        public DeleteControlIf(Wd.ContentControl control) : base(control) { }

        public static bool Test(Wd.ContentControl control)
        {
            return control.PrecedentInstruction() != null &&
                control.PrecedentInstruction().Command == "DeleteControlIf";
        }

        public override void Process()
        {
            var result = PrecedentExpression1.Resolve(control.PrecedentInstruction().Expression, control.PrecedentInstruction().ExpressionData(control.Range.Document));
            if (result)
            {
                var range = control.Range;
                control.Delete(true);

                RemovePreceedingComma(range);
                RemovePreceedingBracket(range);
                RemoveTrailingBracket(range);
                RemoveDoubleSpaces(range);
            }
        }

        static void RemovePreceedingComma(Wd.Range range)
        {
            range.MoveStart(Wd.WdUnits.wdCharacter, -2);
            if (range.Characters.Count == 2 && range.Text == ", ")
                range.Delete();
            range.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);
        }
        static void RemovePreceedingBracket(Wd.Range range)
        {
            range.MoveStart(Wd.WdUnits.wdCharacter, -2);
            if (range.Characters.Count == 2 && range.Text == "(\"")
                range.Delete();
            range.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);
        }
        static void RemoveTrailingBracket(Wd.Range range)
        {
            range.MoveEnd(Wd.WdUnits.wdCharacter, 2);
            if (range.Characters.Count == 2 && range.Text == "\")")
                range.Delete();
            range.Collapse(Wd.WdCollapseDirection.wdCollapseStart);
        }
        static void RemoveDoubleSpaces(Microsoft.Office.Interop.Word.Range range)
        {
            range.MoveStart(Wd.WdUnits.wdCharacter, -1);
            range.MoveEnd(Wd.WdUnits.wdCharacter, 1);
            if (range.Characters.Count == 2 && range.Text == "  ")
                range.Characters.First.Delete();
        }
    }
}
