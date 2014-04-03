using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace GorillaDocs
{
    public class ClipboardHelper
    {
        [DllImport("user32.dll", EntryPoint = "OpenClipboard", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        static extern bool OpenClipboard(uint hWnd);

        [DllImport("user32.dll", EntryPoint = "EmptyClipboard", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        static extern bool EmptyClipboard();

        [DllImport("user32.dll", EntryPoint = "CloseClipboard", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        static extern bool CloseClipboard();

        public static void Clear()
        {
            OpenClipboard(0);
            EmptyClipboard();
            CloseClipboard();
        }
    }
}
