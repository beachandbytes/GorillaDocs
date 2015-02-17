using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace GorillaDocs
{
    public enum OfficeFileType { Word, Excel, PowerPoint, None }

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

        public static List<FileInfo> GetFilesList(this DirectoryInfo folder, string searchPattern, SearchOption searchOption = SearchOption.AllDirectories)
        {
            return folder.GetFiles(searchPattern, searchOption)
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
        public static bool IsExcel(this FileInfo file)
        {
            return (file.Extension.ToLower() == ".xls" || file.Extension.ToLower() == ".xlsx" || file.Extension.ToLower() == ".xlsm" ||
                file.Extension.ToLower() == ".xlt" || file.Extension.ToLower() == ".xltx" || file.Extension.ToLower() == ".xltm");
        }

        public static OfficeFileType Type(this FileInfo file)
        {
            if (file.IsWord())
                return OfficeFileType.Word;
            else if (file.IsExcel())
                return OfficeFileType.Excel;
            else if (file.IsPowerPoint())
                return OfficeFileType.PowerPoint;
            else
                return OfficeFileType.None;
        }
    }
}
