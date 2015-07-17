using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;
using GorillaDocs.Word.Precedent.Delivery;

namespace GorillaDocs.Word.Precedent.Controls
{
    public class ContactDelivery : PrecedentControl
    {
        public ContactDelivery(Wd.ContentControl control) : base(control) { }

        public static bool Test(Wd.ContentControl control)
        {
            return control.GetPrecedentInstruction() != null &&
                control.GetPrecedentInstruction().Command == "ContactDelivery";
        }

        public override void Process()
        {
            var result = PrecedentExpression1.Resolve(control.GetPrecedentInstruction().Expression, control.GetPrecedentInstruction().ExpressionData(control.Range.Document));
            if (result)
                control.DeleteLine();
            else
                control.AsDelivery1().Update();
        }
    }
}
