using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GorillaDocs.SharePoint
{
    public class SPHelper
    {
        public delegate void GetLibrariesSuccessCallback(List<SPLibrary> libraries);
        public delegate void GetFilesSuccessCallback(List<SPFile> files);
        public delegate void FailureCallback(AggregateException ae);

        public static List<SPLibrary> GetLibraries(string webUrl)
        {
            var context = new ClientContext(webUrl);
            var web = context.Web;
            context.Load(web, w => w.Title, w => w.Description);
            var query = from list in web.Lists.Include(l => l.Title, l => l.Id)
                        where list.Hidden == false && list.BaseType == BaseType.DocumentLibrary
                        select list;
            var lists = context.LoadQuery(query);
            context.ExecuteQuery();

            var libraries = new List<SPLibrary>();
            foreach (var list in lists)
                libraries.Add(new SPLibrary()
                {
                    Title = list.Title,
                    Id = list.Id,
                    WebUrl = webUrl
                });
            return libraries;
        }

        public static SPLibrary GetLibrary(string webUrl, string Title)
        {
            var context = new ClientContext(webUrl);
            var web = context.Web;
            context.Load(web, w => w.Title, w => w.Description);
            var query = from list in web.Lists.Include(l => l.Title, l => l.Id)
                        where list.Hidden == false
                            && list.BaseType == BaseType.DocumentLibrary
                            && list.Title == Title
                        select list;
            var lists = context.LoadQuery(query);
            context.ExecuteQuery();

            if (lists.Count() == 1)
                return new SPLibrary()
                {
                    Title = lists.First().Title,
                    Id = lists.First().Id,
                    WebUrl = webUrl
                };
            else
                return null;
        }

        public static void GetLibraries_Async(string webUrl, GetLibrariesSuccessCallback SuccessCallback, FailureCallback FailureCallback)
        {
            Task<List<SPLibrary>> T = Task.Factory.StartNew(() =>
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
            context.Load(files, fs => fs.Include(f => f.Name, f => f.ETag, f => f.ServerRelativeUrl, f => f.ListItemAllFields));
            context.ExecuteQuery();

            var items = new List<SPFile>();
            foreach (File file in files)
                if (extensions == null || extensions.Any(ext => file.Name.EndsWith(ext)))
                {
                    var listitem = file.ListItemAllFields;
                    items.Add(new SPFile()
                    {
                        Name = file.Name.Substring(0, file.Name.LastIndexOf('.')),
                        Extension = file.Name.Substring(file.Name.LastIndexOf('.')),
                        ETag = file.ETag,
                        RemoteUrl = webUrl + file.ServerRelativeUrl,
                        Category = Convert.ToString(listitem.FieldValues["Category"])
                    });
                }
            return new List<SPFile>(items.OrderBy(f => f.Name));
        }

        public static void GetFiles_Async(string webUrl, string libraryTitle, GetFilesSuccessCallback SuccessCallback, FailureCallback FailureCallback)
        {
            Task<List<SPFile>> T = Task.Factory.StartNew(() =>
            {
                return GetFiles(webUrl, libraryTitle, null);
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

        public static TermStoreCollection GetTaxonomyTermStores(string webUrl)
        {
            var context = new ClientContext(webUrl);
            var termStores = TaxonomySession.GetTaxonomySession(context).TermStores;
            context.Load(termStores);
            context.ExecuteQuery();
            return termStores;
        }

        public static TermGroupCollection GetTaxonomyTermGroups(string webUrl, Guid termStoreId)
        {
            var context = new ClientContext(webUrl);
            var termStores = TaxonomySession.GetTaxonomySession(context).TermStores;
            context.Load(termStores);
            context.ExecuteQuery();
            var termStore = termStores.Where(t => t.Id == termStoreId).FirstOrDefault();
            var termGroups = termStore.Groups;
            context.Load(termGroups);
            context.ExecuteQuery();
            return termGroups;
        }

        public static TermSetCollection GetTaxonomyTermSets(string webUrl, Guid termStoreId, Guid groupId)
        {
            var context = new ClientContext(webUrl);
            var termStores = TaxonomySession.GetTaxonomySession(context).TermStores;
            context.Load(termStores);
            context.ExecuteQuery();
            var termStore = termStores.Where(t => t.Id == termStoreId).FirstOrDefault();
            var termGroups = termStore.Groups;
            context.Load(termGroups);
            context.ExecuteQuery();
            var termGroup = termGroups.Where(t => t.Id == groupId).FirstOrDefault();
            var termSets = termGroup.TermSets;
            context.Load(termSets);
            context.ExecuteQuery();
            return termSets;
        }

    }
}
