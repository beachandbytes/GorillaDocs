using AddinExpress.Outlook;
using GorillaDocs.libs.PostSharp;

namespace GorillaDocs
{
    [Log]
    /// <summary>
    /// Note: I don't dispose the security manager because it looses it's registration somehow..
    /// </summary>
    public static class OutlookSecurityManager
    {
        public static SecurityManager securityManager = new SecurityManager();
    }
}

