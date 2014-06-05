using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;

namespace GorillaDocs.SharePoint
{
    public static class SPFileHelper
    {
        public static void Download(this SPFile file, DirectoryInfo folder)
        {
            using (var client = new WebClient())
            {
                client.UseDefaultCredentials = true;
                file.SetLocalFullName(folder);
                client.DownloadFile(file.RemoteUrl, file.LocalFullName);
            }
        }

        public static void SetLocalFullName(this SPFile file, DirectoryInfo folder)
        {
            file.LocalFullName = String.Format("{0}\\{1}{2}", folder.FullName, file.Name, file.Extension);
        }

        public static bool ExtensionMatches(this SPFile file, string RegexPattern)
        {
            return Regex.IsMatch(file.Extension, RegexPattern);
        }
    }
}
