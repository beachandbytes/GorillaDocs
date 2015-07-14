using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.Precedent.Controls
{
    public abstract class PrecedentControl
    {
        protected readonly Wd.ContentControl control;
        protected readonly Wd.Document doc;

        public PrecedentControl(Wd.ContentControl control)
        {
            this.control = control;
            this.doc = control.Range.Document;
        }

        public abstract void Process();
    }
}
