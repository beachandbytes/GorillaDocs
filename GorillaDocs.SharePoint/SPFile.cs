using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs.SharePoint
{
    public class SPFile
    {
        public string Name { get; set; }
        public string Extension { get; set; }
        public string ETag { get; set; }
        public string RemoteUrl { get; set; } // Used string instead of Uri because Serialization breaks
        public string LocalFullName { get; set; }
        public string Category { get; set; }
    }
}
