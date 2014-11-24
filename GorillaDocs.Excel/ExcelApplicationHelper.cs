using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using XL = Microsoft.Office.Interop.Excel;

namespace GorillaDocs.Excel
{
    public class ExcelApplicationHelper
    {
        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);
        const string ProcessName = "EXCEL";

        public static XL.Application GetApplication()
        {
            try
            {
                if (!IsExcelRunning())
                    return CreateExcelApp();

                int i = 0;
                while (true)
                {
                    try
                    {
                        return (XL.Application)Marshal.GetActiveObject("Excel.Application");
                    }
                    catch
                    {
                        if (i > 10)
                            throw new InvalidOperationException("Unable to start Excel Application.");
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

        static XL.Application CreateExcelApp()
        {
            var app = new XL.Application() { Visible = true };
            app.Activate(true);
            return app;
        }

        static bool IsExcelRunning() { return Process.GetProcessesByName(ProcessName).Any(); }
    }

    public static class ExcelApplicationExtensionMethods
    {
        [System.Diagnostics.DebuggerStepThrough]
        public static void Activate(this XL.Application app, bool WaitIfBusy = false, int RetryAttempt = 0)
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
