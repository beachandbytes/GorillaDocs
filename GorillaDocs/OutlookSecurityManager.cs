using AddinExpress.Outlook;

namespace GorillaDocs
{
    /// <summary>
    /// Note: I don't dispose the security manager because it looses it's registration somehow..
    /// </summary>
    public static class OutlookSecurityManager
    {
        public static SecurityManager securityManager = new SecurityManager();
    }
}

