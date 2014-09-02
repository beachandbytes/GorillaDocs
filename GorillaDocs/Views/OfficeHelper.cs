using GorillaDocs.libs.PostSharp;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace GorillaDocs.Views
{
    [Log]
    public class OfficeDialog : Window
    {
        [DllImport("user32.dll")]
        static extern int GetWindowLong(IntPtr hwnd, int index);
        [DllImport("user32.dll")]
        static extern int SetWindowLong(IntPtr hwnd, int index, int newStyle);
        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hwnd, IntPtr hwndInsertAfter, int x, int y, int width, int height, uint flags);
        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);

        const int GWL_EXSTYLE = -20;
        const int WS_EX_DLGMODALFRAME = 0x0001;
        const int SWP_NOSIZE = 0x0001;
        const int SWP_NOMOVE = 0x0002;
        const int SWP_NOZORDER = 0x0004;
        const int SWP_FRAMECHANGED = 0x0020;
        const uint WM_SETICON = 0x0080;
        const int ICON_SMALL = 0;
        const int ICON_BIG = 1;

        /// <summary>
        /// Sometimes get System.ComponentModel.Win32Exception: Invalid window handle
        /// I'm pretty sure that this is because Word is shit at handling windows and has an internal memory leak
        /// http://stackoverflow.com/questions/222649/winforms-issue-error-creating-window-handle
        /// I'm not sure why this error isn't trapped and logged by the try catch below. Somehow it bubbles up to the calling routine..
        /// </summary>
        public OfficeDialog()
        {
            this.ShowInTaskbar = false;
            //this.Topmost = true;

            //Uri uri = new Uri("PresentationFramework.Aero;V3.0.0.0;31bf3856ad364e35;component\\themes/aero.normalcolor.xaml", UriKind.Relative);
            //Uri uri = new Uri("PresentationFramework.Classic;V3.0.0.0;31bf3856ad364e35;component\\themes/classic.xaml", UriKind.Relative);
            //Resources.MergedDictionaries.Add(Application.LoadComponent(uri) as ResourceDictionary);

            //var helper = new WindowInteropHelper(this);
            //using (Process currentProcess = Process.GetCurrentProcess())
            //    helper.Owner = currentProcess.MainWindowHandle;
        }

        public new void ShowDialog()
        {
            try
            {
                var helper = new WindowInteropHelper(this);
                using (Process currentProcess = Process.GetCurrentProcess())
                    helper.Owner = currentProcess.MainWindowHandle;
                base.ShowDialog();
            }
            catch (System.ComponentModel.Win32Exception ex)
            {
                Message.LogWarning(ex);
                //this.Topmost = true;
                var helper = new WindowInteropHelper(this);
                using (Process currentProcess = Process.GetCurrentProcess())
                    helper.Owner = currentProcess.MainWindowHandle;
                base.ShowDialog();
            }
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            RemoveIcon(this);
            HideMinimizeAndMaximizeButtons(this);
            //using (Process currentProcess = Process.GetCurrentProcess())
            //    SetCentering(this, currentProcess.MainWindowHandle);
        }

        public static void HideMinimizeAndMaximizeButtons(Window window)
        {
            const int GWL_STYLE = -16;

            IntPtr hwnd = new WindowInteropHelper(window).Handle;
            long value = GetWindowLong(hwnd, GWL_STYLE);

            SetWindowLong(hwnd, GWL_STYLE, (int)(value & -131073 & -65537));
        }

        public static void RemoveIcon(Window w)
        {
            // Get this window's handle 
            IntPtr hwnd = new WindowInteropHelper(w).Handle;

            // Change the extended window style to not show a window icon
            int extendedStyle = OfficeDialog.GetWindowLong(hwnd, GWL_EXSTYLE);
            OfficeDialog.SetWindowLong(hwnd, GWL_EXSTYLE, extendedStyle | WS_EX_DLGMODALFRAME);

            // reset the icon, both calls important
            OfficeDialog.SendMessage(hwnd, WM_SETICON, (IntPtr)ICON_SMALL, IntPtr.Zero);
            OfficeDialog.SendMessage(hwnd, WM_SETICON, (IntPtr)ICON_BIG, IntPtr.Zero);

            // Update the window's non-client area to reflect the changes
            OfficeDialog.SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
        }

        static void SetCentering(Window win, IntPtr ownerHandle)
        {
            bool isWindow = IsWindow(ownerHandle);
            if (!isWindow) //Don't try and centre the window if the ownerHandle is invalid.  To resolve issue with invalid window handle error
            {
                //Message.LogInfo(string.Format("ownerHandle IsWindow: {0}", isWindow));
                return;
            }
            //Show in center of owner if win form.
            if (ownerHandle.ToInt32() != 0)
            {
                var helper = new WindowInteropHelper(win);
                helper.Owner = ownerHandle;
                win.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }
            else
                win.WindowStartupLocation = WindowStartupLocation.CenterOwner;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsWindow(IntPtr hWnd);
   }
}
