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
                control.GetPrecedentInstruction() != null &&
                control.GetPrecedentInstruction().Command == "DeleteColumnIf";
        }

        public override void Process()
        {
            var result = PrecedentExpression1.Resolve(control.GetPrecedentInstruction().Expression, control.GetPrecedentInstruction().ExpressionData(control.Range.Document));
            if (result)
                control.DeleteColumnAndAutoFit();
            else
                control.Delete();
        }
    }
}
