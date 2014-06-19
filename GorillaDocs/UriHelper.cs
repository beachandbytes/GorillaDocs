using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs
{
    public static class UriHelper
    {
        public static string Server(this Uri uri)
        {
            return String.Format(@"{0}://{1}", uri.Scheme, uri.Host);
        }
    }
}
