using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GorillaDocs.SharePoint
{
    public class SPHelper
    {
        public delegate void GetLibrariesSuccessCallback(List<string> libraries);
        public delegate void GetFilesSuccessCallback(List<SPFile> files);
        public delegate void FailureCallback(AggregateException ae);

        public static List<string> GetLibraries(string webUrl)
        {
            var context = new ClientContext(webUrl);
            var web = context.Web;
            context.Load(web, w => w.Title, w => w.Description);
            var query = from list in web.Lists.Include(l => l.Title)
                        where list.Hidden == false && list.BaseType == BaseType.DocumentLibrary
                        select list;
            var lists = context.LoadQuery(query);
            context.ExecuteQuery();
            var libraries = new List<string>();
            foreach (var list in lists)
                libraries.Add(list.Title);
            return libraries;
        }

        public static void GetLibraries_Async(string webUrl, GetLibrariesSuccessCallback SuccessCallback, FailureCallback FailureCallback)
        {
            Task<List<string>> T = Task.Factory.StartNew(() =>
                {
                    return GetLibraries(webUrl);
                });

            T.ContinueWith((antecedent) =>
                {
                    try
                    {
                        SuccessCallback(antecedent.Result);
                    }
                    catch (AggregateException ae)
                    {
                        FailureCallback(ae);
                    }
                });
        }

        public static List<SPFile> GetFiles(string webUrl, string libraryTitle, String[] extensions = null)
        {
            var context = new ClientContext(webUrl);
            var web = context.Web;

            var list = web.Lists.GetByTitle(libraryTitle);
            var files = list.RootFolder.Files;
            context.Load(files, fs => fs.Include(f => f.Name));
            context.ExecuteQuery();

            var items = new List<SPFile>();
            foreach (File file in files)
                if (extensions == null || extensions.Any(ext => file.Name.EndsWith(ext)))
                    items.Add(new SPFile()
                    {
                        Name = file.Name.Substring(0, file.Name.LastIndexOf('.')),
                        Extension = file.Name.Substring(file.Name.LastIndexOf('.')),
                        HashCode = file.GetHashCode()
                    });
            return items;
        }

        public static void GetFiles_Async(string webUrl, string libraryTitle, GetFilesSuccessCallback SuccessCallback, FailureCallback FailureCallback)
        {
            Task<List<SPFile>> T = Task.Factory.StartNew(() =>
            {
                return GetFiles(webUrl, libraryTitle);
            });

            T.ContinueWith((antecedent) =>
            {
                try
                {
                    SuccessCallback(antecedent.Result);
                }
                catch (AggregateException ae)
                {
                    FailureCallback(ae);
                }
            });
        }
    }
}
