using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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
            UriBuilder uri = new UriBuilder(assembly.CodeBase);
            string path = Uri.UnescapeDataString(uri.Path);
            return System.IO.Path.GetDirectoryName(path);
        }

        public static DirectoryInfo Folder(this Assembly assembly) { return new DirectoryInfo(assembly.Path()); }

        public static string FullPath(this Assembly assembly)
        {
            UriBuilder uri = new UriBuilder(assembly.CodeBase);
            string path = Uri.UnescapeDataString(uri.Path);
            return String.Format("{0}\\{1}", System.IO.Path.GetDirectoryName(path), assembly.ManifestModule.Name);
        }
    }
}
