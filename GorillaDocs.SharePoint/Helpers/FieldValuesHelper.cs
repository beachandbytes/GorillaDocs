using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs.SharePoint
{
    public static class FieldValuesHelper
    {
        public static string GetValue(this Dictionary<string,object> FieldValues, string key)
        {
            return FieldValues.Any(x => x.Key == key) ? Convert.ToString(FieldValues[key]) : "";
        }
    }
}
