using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.Precedent.Controls
{
    class DeleteColumnIf : PrecedentControl
    {
        public DeleteColumnIf(Wd.ContentControl control) : base(control) { }

        public static bool Test(Wd.ContentControl control)
        {
            return control.Type == Wd.WdContentControlType.wdContentControlRichText &&
                control.PrecedentInstruction() != null &&
                control.PrecedentInstruction().Command == "DeleteColumnIf";
        }

        public override void Process()
        {
            var result = PrecedentExpression1.Resolve(control.PrecedentInstruction().Expression, control.PrecedentInstruction().ExpressionData(control.Range.Document));
            if (result)
                control.DeleteColumnAndAutoFit();
            else
                control.Delete();
        }
    }
}
