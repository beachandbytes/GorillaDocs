﻿using GorillaDocs.libs.PostSharp;
using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    [Log]
    public class WordApplicationHelper
    {
        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);
        const string ProcessName = "WINWORD";

        public static Wd.Application GetApplication()
        {
            try
            {
                if (!IsWordRunning())
                    return CreateWordApp();

                int i = 0;
                while (true)
                {
                    try
                    {
                        return (Wd.Application)Marshal.GetActiveObject("Word.Application");
                    }
                    catch
                    {
                        if (i > 10)
                            throw new InvalidOperationException("Unable to start Word Application.");
                        i++;
                        Thread.Sleep(1000);
                    }
                }
            }
            finally
            {
                SetForegroundWindow(Process.GetProcessesByName(ProcessName).First().MainWindowHandle);
            }
        }

        static Wd.Application CreateWordApp()
        {
            var app = new Wd.Application() { Visible = true };
            app.Activate(true);
            return app;
        }

        static bool IsWordRunning() { return Process.GetProcessesByName(ProcessName).Any(); }
    }

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
                        Thread.Sleep(2000);
                        app.Activate(WaitIfBusy, ++RetryAttempt);
                    }
                    else
                        throw ex;
            }
        }
    }
}
