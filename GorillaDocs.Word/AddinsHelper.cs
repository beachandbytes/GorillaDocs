using GorillaDocs.libs.PostSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using O = Microsoft.Office.Core;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    [Log]
    public static class AddinsHelper
    {
        public static bool IsLoaded(this O.COMAddIns commAddins, string name)
        {
            O.COMAddIn item = commAddins.Find(name);
            if (item != null && item.Connect)
                return true;
            return false;
        }

        public static bool Exists(this O.COMAddIns comAddins, string name) { return comAddins.Find(name) != null; }

        public static O.COMAddIn Find(this O.COMAddIns comAddins, string name)
        {
            foreach (O.COMAddIn item in comAddins)
                if (item.ProgId.Contains(name, StringComparison.OrdinalIgnoreCase))
                    return item;
            return null;
        }

        public static void Disable(this Wd.AddIns addins, string name)
        {
            foreach (Wd.AddIn addin in addins)
                if (addin.Name == name)
                    addin.Installed = false;
        }

        public static void Enable(this Wd.AddIns addins, string name)
        {
            foreach (Wd.AddIn addin in addins)
                if (addin.Name == name)
                    addin.Installed = true;
        }

        public static void Enable(this Wd.AddIns addins, FileInfo file)
        {
            foreach (Wd.AddIn addin in addins)
                if (addin.Name == file.Name)
                {
                    addin.Installed = true;
                    return;
                }
            addins.Add(file.FullName);
        }

        public static bool Exists(this Wd.AddIns addins, string name)
        {
            foreach (Wd.AddIn addin in addins)
                if (addin.Name == name)
                    return true;
            return false;
        }

    }
}
