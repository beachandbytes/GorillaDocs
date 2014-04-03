using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

namespace GorillaDocs
{
    public static class AssemblyHelper
    {
        public static string Title(this Assembly assembly)
        {
            return ((AssemblyTitleAttribute)assembly.GetCustomAttributes(typeof(AssemblyTitleAttribute), false)[0]).Title;
        }

        public static string FileVersion(this Assembly assembly)
        {
            return FileVersionInfo.GetVersionInfo(assembly.Location).FileVersion;
        }

        public static string Path(this Assembly assembly)
        {
            Uri uri = new Uri(assembly.CodeBase);
            string path = uri.LocalPath;
            return path.Substring(0, path.LastIndexOf("\\"));
        }

    }
}
