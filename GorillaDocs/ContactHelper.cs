using GorillaDocs.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs
{
    public static class ContactHelper
    {
        public static bool IsNullOrEmpty(this Contact contact) { return contact == null || contact.IsEmpty(); }
    }
}
