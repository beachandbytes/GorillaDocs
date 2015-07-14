using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Threading;

namespace GorillaDocs
{
    public static class DispatcherHelper
    {
        public static void WaitUntilApplicationIdle(this Dispatcher dispatcher, Action action)
        {
            dispatcher.Invoke(action, DispatcherPriority.ApplicationIdle);
        }
    }
}
