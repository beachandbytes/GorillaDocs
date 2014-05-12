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

        public static List<int> ToIntList(this string value, char split)
        {
            string[] parts = value.Split(split);
            List<int> list = new List<int>();
            foreach (string part in parts)
                list.Add(int.Parse(part));
            return list;
        }

        public static int ToVal(this string value)
        {
            int returnValue;
            if (int.TryParse(value, out returnValue))
                return returnValue;
            else
                return 0;
        }
    }
}
