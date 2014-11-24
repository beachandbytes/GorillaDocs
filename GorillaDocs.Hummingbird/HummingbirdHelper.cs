using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs.Hummingbird
{
    public class HummingbirdHelper
    {
        public static bool IsInstalled()
        {
            using (RegistryKey reg = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Wow6432Node\\Hummingbird"))
                return reg != null;
        }
    }
}
