using System.IO;
using System.Reflection;

namespace GorillaDocs
{
    public class AppConfig
    {
        public static FileInfo GetFile(Assembly assembly = null)
        {
            if (assembly == null)
                assembly = Assembly.GetCallingAssembly();
            string path = ConvertFromFileProtocol(assembly.CodeBase);
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
