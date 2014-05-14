using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace GorillaDocs
{
    public static class DateTimeHelper
    {
        public static DateTime? AsNullableDateTime(this string value, CultureInfo culture)
        {
            DateTime result;
            if (DateTime.TryParse(value, out result))
                return result;
            else if (DateTime.TryParse(value, culture, DateTimeStyles.None, out result))
                return result;
            else
                throw new InvalidOperationException(string.Format("Unable to convert '{0}' to DateTime value.", value));
        }
    }
}
