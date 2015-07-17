using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.Precedent.Controls
{
    class DeleteRowIf_OrRemoveControl : PrecedentControl
    {
        public DeleteRowIf_OrRemoveControl(Wd.ContentControl control) : base(control) { }

        public static bool Test(Wd.ContentControl control)
        {
            return control.GetPrecedentInstruction() != null &&
                control.GetPrecedentInstruction().Command == "DeleteRowIf_OrRemoveControl";
        }

        public override void Process()
        {
            var result = PrecedentExpression1.Resolve(control.GetPrecedentInstruction().Expression, control.GetPrecedentInstruction().ExpressionData(control.Range.Document));
            if (result)
                control.DeleteRow();
            else 
                control.Delete();
        }
    }
}
