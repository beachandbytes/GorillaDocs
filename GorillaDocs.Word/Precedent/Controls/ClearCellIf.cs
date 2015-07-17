using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.Precedent.Controls
{
    public class ClearCellIf : PrecedentControl
    {
        public ClearCellIf(Wd.ContentControl control) : base(control) { }

        public static bool Test(Wd.ContentControl control)
        {
            return control.Type == Wd.WdContentControlType.wdContentControlRichText &&
                control.GetPrecedentInstruction() != null &&
                control.GetPrecedentInstruction().Command == "ClearCellIf";
        }

        public override void Process()
        {
            var result = PrecedentExpression1.Resolve(control.GetPrecedentInstruction().Expression, control.GetPrecedentInstruction().ExpressionData(control.Range.Document));
            if (result)
                control.ClearCell();
            else
                control.Delete();
        }
    }
}
