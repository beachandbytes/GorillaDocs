using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using PP = Microsoft.Office.Interop.PowerPoint;
using O = Microsoft.Office.Core;

namespace GorillaDocs.PowerPoint
{
    public class PowerPointApplicationHelper
    {
        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);
        const string ProcessName = "POWERPNT";

        public static PP.Application GetApplication()
        {
            try
            {
                if (!IsPowerPointRunning())
                    return CreatePowerPointApp();

                int i = 0;
                while (true)
                {
                    try
                    {
                        return (PP.Application)Marshal.GetActiveObject("PowerPoint.Application");
                    }
                    catch
                    {
                        if (i > 10)
                            throw new InvalidOperationException("Unable to start PowerPoint Application.");
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

        static PP.Application CreatePowerPointApp()
        {
            var app = new PP.Application() { Visible = O.MsoTriState.msoTrue };
            app.Activate(true);
            return app;
        }

        static bool IsPowerPointRunning() { return Process.GetProcessesByName(ProcessName).Any(); }
    }

    public static class PowerPointApplicationExtensionMethods
    {
        [System.Diagnostics.DebuggerStepThrough]
        public static void Activate(this PP.Application app, bool WaitIfBusy = false, int RetryAttempt = 0)
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
