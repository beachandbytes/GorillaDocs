using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs.SharePoint
{
    public class CheckInHelper
    {
        public static void DiscardCheckOut(string webUrl, string fileUrl)
        {
            Uri fileUri = new Uri(fileUrl);
            string server = fileUri.AbsoluteUri.Replace(fileUri.AbsolutePath, "");
            var context = new ClientContext(fileUrl);
            Uri fileSiteUri = Web.WebUrlFromPageUrlDirect(context, fileUri);
            
            //var context = new ClientContext(webUrl);
            var web = context.Web;

            var uri = new Uri(fileUrl);
            var file = web.GetFileByServerRelativeUrl(uri.LocalPath);
            file.UndoCheckOut();
            context.Load(file);
            context.ExecuteQuery();
        }
    }
}
