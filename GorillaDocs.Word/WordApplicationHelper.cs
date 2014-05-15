using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public class WordApplicationHelper
    {
        public static Wd.Application GetWordApplication()
        {
            if (!IsWordRunning())
                return CreateWordApp();

            int i = 0;
            while (true)
            {
                try
                {
                    return (Wd.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                }
                catch
                {
                    if (i > 10)
                        throw new InvalidOperationException("Unable to start Word Application.");
                    i++;
                    System.Threading.Thread.Sleep(1000);
                }
            }
        }

        static Wd.Application CreateWordApp()
        {
            var wordApp = new Wd.Application() { Visible = true };
            wordApp.Activate();
            return wordApp;
        }

        static bool IsWordRunning()
        {
            foreach (Process x in Process.GetProcesses())
                if (x.ProcessName.Contains("WINWORD"))
                    return true;
            return false;
        }

    }
}
