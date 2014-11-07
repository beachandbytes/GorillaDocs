using System;
using System.Diagnostics;

namespace GorillaDocs
{
    public class Pdf
    {
        public static void Open(string path, string namedDestination = "")
        {
            try
            {
                //Process.Start("AcroRd32.exe", String.Format(" /n /A \"pagemode=bookmarks&nameddest={0}\" \"{1}\"", namedDestination, path));
                Process.Start("AcroRd32.exe", String.Format(" /n /A \"page={0}\" \"{1}\"", namedDestination, path));
            }
            catch
            {
                Process.Start(path);
            }
        }
    }
}
