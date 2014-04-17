using System;
using System.Collections.Generic;
using System.Linq;
using O = Microsoft.Office.Core;

namespace GorillaDocs.Word
{
    public static class COMAddinsHelper
    {
        public static bool IsLoaded(this O.COMAddIns commAddins, string name)
        {
            O.COMAddIn item = commAddins.Find(name);
            if (item != null && item.Connect)
                return true;
            return false;
        }

        public static bool Exists(this O.COMAddIns comAddins, string name)
        {
            return comAddins.Find(name) != null;
        }

        public static O.COMAddIn Find(this O.COMAddIns comAddins, string name)
        {
            foreach (O.COMAddIn item in comAddins)
                if (item.ProgId.Contains(name, StringComparison.OrdinalIgnoreCase))
                    return item;
            return null;
        }
    }
}
