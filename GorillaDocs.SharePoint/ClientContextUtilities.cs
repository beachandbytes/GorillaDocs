using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;

namespace GorillaDocs.SharePoint
{
    public static class ClientContextUtilities
    {
        /// <summary>
        /// Resolve client context  
        /// </summary>
        /// <param name="requestUri"></param>
        /// <param name="context"></param>
        /// <param name="credentials"></param>
        /// <returns></returns>
        [System.Diagnostics.DebuggerStepThrough]
        public static bool TryResolveClientContext(Uri requestUri, out ClientContext context, ICredentials credentials = null)
        {
            context = null;
            var baseUrl = requestUri.GetLeftPart(UriPartial.Authority);
            for (int i = requestUri.Segments.Length; i >= 0; i--)
            {
                var path = string.Join(string.Empty, requestUri.Segments.Take(i));
                string url = string.Format("{0}{1}", baseUrl, path);
                try
                {
                    context = new ClientContext(url);
                    if (credentials != null)
                        context.Credentials = credentials;
                    context.ExecuteQuery();
                    return true;
                }
                catch { }
            }
            return false;
        }
    }
}
