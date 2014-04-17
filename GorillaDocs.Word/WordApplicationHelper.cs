using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public class WordApplicationHelper
    {
        public static Wd.Application CreateWordApp()
        {
            var wordApp = new Wd.Application() { Visible = true };
            wordApp.Activate();
            return wordApp;
        }
    }
}
