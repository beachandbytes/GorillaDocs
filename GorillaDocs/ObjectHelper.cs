using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs
{
    // No LOG Attribute - Can not log because this code is called by the logging code
    public static class ObjectHelper
    {
        public static Type NullableGetType(this Object obj)
        {
            if (obj == null)
                return null;
            else
                return obj.GetType();
        }
    }
}
