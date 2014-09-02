using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class DialogHelper
    {
        public static void SetName(this Wd.Dialog dialog, string name)
        {
            dialog.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, (object)dialog, new object[] { name });
        }
        public static string GetName(this Wd.Dialog dialog)
        {
            return Convert.ToString(dialog.GetType().InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, dialog, null));
        }
    }
}
