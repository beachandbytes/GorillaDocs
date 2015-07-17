using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.ControlManagers
{
    public class ControlFragment
    {
        public string Title { get; set; }
        public bool? Replace { get; set; }
        public string[] AdditionalControlsTitles { get; set; }
    }
}
