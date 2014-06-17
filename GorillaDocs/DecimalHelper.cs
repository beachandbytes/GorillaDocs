using GorillaDocs.libs.PostSharp;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace GorillaDocs
{
    [Log]
    public static class DecimalHelper
    {
        public static string AsDecimalString(this string value, CultureInfo culture)
        {
            return string.Format(culture, "{0:N}", value.ToDecimal(culture));
        }

        public static decimal ToDecimal(this string value, CultureInfo culture)
        {
            decimal result = 0;
            if (decimal.TryParse(value, NumberStyles.Any, culture, out result))
                return result;
            else
                throw new ArgumentException(string.Format("'{0}' is not a decimal", value));
        }

        public static bool IsDecimal(this string value, CultureInfo culture)
        {
            decimal result = 0;
            return decimal.TryParse(value, NumberStyles.Any, culture, out result);
        }
    }
}
