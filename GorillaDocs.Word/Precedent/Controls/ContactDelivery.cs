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
            return control.PrecedentInstruction() != null &&
                control.PrecedentInstruction().Command == "ContactDelivery";
        }

        public override void Process()
        {
            control.AsDelivery1().Update();
        }
    }
}
