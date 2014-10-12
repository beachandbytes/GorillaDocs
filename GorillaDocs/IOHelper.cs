using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace GorillaDocs
{
    // No LOG Attribute - Can not log because this code is called by the logging code
    public static class IOHelper
    {
        public static string NameWithoutExtension(this FileInfo file)
        {
            return file.Name.Replace(file.Extension, "");
        }
        public static string Path(this FileInfo file)
        {
            return file.FullName.Substring(0, file.FullName.LastIndexOf('\\'));
        }
        public static bool ContainsFiles(this DirectoryInfo folder, string searchPattern)
        {
            return folder.GetFiles(searchPattern).Where(x => !x.Name.StartsWith("~$")).Count() > 0;
        }

        public static List<FileInfo> GetFilesList(this DirectoryInfo folder, string searchPattern)
        {
            return folder.GetFiles(searchPattern, SearchOption.AllDirectories)
                .Where(x => !x.Name.StartsWith("~$"))
                .OrderBy(x => x.Name)
                .ToList();
        }

        public static bool IsPowerPoint(this FileInfo file)
        {
            return (file.Extension.ToLower() == ".pot" || file.Extension.ToLower() == ".potx" || file.Extension.ToLower() == ".potm" ||
                file.Extension.ToLower() == ".ppt" || file.Extension.ToLower() == ".pptx" || file.Extension.ToLower() == ".pptm" ||
                file.Extension.ToLower() == ".pps" || file.Extension.ToLower() == ".ppsx" || file.Extension.ToLower() == ".ppsm");
        }
        public static bool IsWord(this FileInfo file)
        {
            return (file.Extension.ToLower() == ".doc" || file.Extension.ToLower() == ".docx" || file.Extension.ToLower() == ".docm" ||
                file.Extension.ToLower() == ".dot" || file.Extension.ToLower() == ".dotx" || file.Extension.ToLower() == ".dotm");
        }
    }
}
