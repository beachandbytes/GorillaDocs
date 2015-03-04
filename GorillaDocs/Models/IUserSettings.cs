using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs.Models
{
    public interface IUserSettings
    {
        string Last_Selected_Office { get; set; }
        int Last_Selected_Templates_Tab { get; set; }
        string Last_Selected_Template { get; set; }
    }
}
