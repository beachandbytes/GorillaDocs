using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GorillaDocs
{
    public class StringArray
    {
        public static bool Contains(string[] ControlNames, string x) { return Array.IndexOf(ControlNames, x) >= 0; }
    }
}
