using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs
{
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
