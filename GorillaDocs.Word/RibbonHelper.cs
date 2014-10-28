using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using O = Microsoft.Office.Core;

namespace GorillaDocs.Word
{
    public static class RibbonHelper
    {
        public static bool IsVisible(this O.IRibbonControl control, Func<bool> func)
        {
            try
            {
                return func();
            }
            catch (COMException ex)
            {
                Message.LogWarning(ex);
                return false;
            }
            catch (InvalidOperationException ex)
            {
                Message.LogWarning(ex);
                return false;
            }
            catch (Exception ex)
            {
                Message.LogError(ex);
                return false;
            }
        }
    }
}
