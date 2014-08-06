using GorillaDocs.libs.PostSharp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    [Log]
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
            wordApp.Activate(true);
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

    [Log]
    public static class WordApplicationExtensionMethods
    {
        [System.Diagnostics.DebuggerStepThrough]
        public static void Activate(this Wd.Application app, bool WaitIfBusy = false, int RetryAttempt = 0)
        {
            try
            {
                app.Activate();
            }
            catch (COMException ex)
            {
                if (ex.ErrorCode == -2146823687) // Cannot activate application
                    if (WaitIfBusy && RetryAttempt < 10)
                    {
                        System.Threading.Thread.Sleep(2000);
                        app.Activate(WaitIfBusy, ++RetryAttempt);
                    }
                    else
                        throw ex;
            }
        }
    }
}
