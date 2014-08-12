using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace GorillaDocs
{
    public class AppConfig
    {
        public static FileInfo GetFile()
        {
            string path = ConvertFromFileProtocol(Assembly.GetCallingAssembly().CodeBase);
            path = path + ".config";
            return new FileInfo(path);
        }

        public static string ConvertFromFileProtocol(string path)
        {
            path = path.ToLower();
            path = path.Replace("file:///", "");
            return path.Replace("/", "\\");
        }
    }
}
