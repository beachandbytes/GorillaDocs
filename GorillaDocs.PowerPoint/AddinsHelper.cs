using System;
using System.Collections.Generic;
using System.Linq;
using O = Microsoft.Office.Core;
using PP = Microsoft.Office.Interop.PowerPoint;

namespace GorillaDocs.PowerPoint
{
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
                if (Contains(item.ProgId, name, StringComparison.OrdinalIgnoreCase))
                    return item;
            return null;
        }

        static bool Contains(string source, string toCheck, StringComparison comparison) { return source.IndexOf(toCheck, comparison) >= 0; }

        public static void Disable(this PP.AddIns addins, string name)
        {
            foreach (PP.AddIn addin in addins)
                if (addin.Name == name)
                    addin.Loaded = O.MsoTriState.msoFalse;
        }

        public static void Enable(this PP.AddIns addins, string name)
        {
            foreach (PP.AddIn addin in addins)
                if (addin.Name == name)
                    addin.Loaded = O.MsoTriState.msoTrue;
        }
    }
}
