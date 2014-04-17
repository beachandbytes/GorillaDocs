using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs
{
    public static class StringHelper
    {
        public static string Remove(this string value, string remove)
        {
            return value.Replace(remove, string.Empty);
        }

        public static bool Contains(this string source, string toCheck, StringComparison comparison)
        {
            return source.IndexOf(toCheck, comparison) >= 0;
        }
    }
}
