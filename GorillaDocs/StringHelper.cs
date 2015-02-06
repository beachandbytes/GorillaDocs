using GorillaDocs.libs.PostSharp;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs
{
    [Log]
    public static class StringHelper
    {
        public static string ToRoman(this int number)
        {
            if ((number < 0) || (number > 3999)) throw new ArgumentOutOfRangeException("insert value betwheen 1 and 3999");
            if (number < 1) return string.Empty;
            if (number >= 1000) return "M" + ToRoman(number - 1000);
            if (number >= 900) return "CM" + ToRoman(number - 900); //EDIT: i've typed 400 instead 900
            if (number >= 500) return "D" + ToRoman(number - 500);
            if (number >= 400) return "CD" + ToRoman(number - 400);
            if (number >= 100) return "C" + ToRoman(number - 100);
            if (number >= 90) return "XC" + ToRoman(number - 90);
            if (number >= 50) return "L" + ToRoman(number - 50);
            if (number >= 40) return "XL" + ToRoman(number - 40);
            if (number >= 10) return "X" + ToRoman(number - 10);
            if (number >= 9) return "IX" + ToRoman(number - 9);
            if (number >= 5) return "V" + ToRoman(number - 5);
            if (number >= 4) return "IV" + ToRoman(number - 4);
            if (number >= 1) return "I" + ToRoman(number - 1);
            throw new ArgumentOutOfRangeException("something bad happened");
        }

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

        public static bool IsIn(this string x, string[] ControlNames) { return Array.IndexOf(ControlNames, x) >= 0; }
    }
}
