﻿using System;
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
            return folder.GetFiles(searchPattern).Count() > 0;
        }
    }
}
