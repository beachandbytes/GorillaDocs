using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GorillaDocs.SharePoint
{
    public class TaxonomyHelper
    {
        public static TermStoreCollection GetTermStores(string webUrl)
        {
            var context = new ClientContext(webUrl);
            var termStores = TaxonomySession.GetTaxonomySession(context).TermStores;
            context.Load(termStores);
            context.ExecuteQuery();
            return termStores;
        }

        public static TermGroupCollection GetTermGroups(string webUrl, Guid termStoreId)
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

        public static TermSetCollection GetTermSets(string webUrl, Guid termStoreId, Guid groupId)
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

        public static Guid GetTermId(Uri requestUri, string TermLabel)
        {
            ClientContext context;
            if (ClientContextUtilities.TryResolveClientContext(requestUri, out context, null))
            {
                using (context)
                {
                    var termStores = TaxonomySession.GetTaxonomySession(context).TermStores;
                    context.Load(termStores);
                    context.ExecuteQuery();
                    foreach (TermStore termStore in termStores)
                    {
                        var labelMatchInfo = new LabelMatchInformation(context)
                        {
                            TermLabel = TermLabel,
                            DefaultLabelOnly = true,
                            StringMatchOption = StringMatchOption.StartsWith,
                            ResultCollectionSize = 100,
                            TrimUnavailable = true
                        };
                        var terms = termStore.GetTerms(labelMatchInfo);
                        context.Load(terms);
                        context.ExecuteQuery();
                        var term = terms.FirstOrDefault();
                        if (term != null)
                            return term.Id;
                    }
                }
            }
            throw new InvalidOperationException(string.Format("Unable to find term '{0}' at '{1}'", TermLabel, requestUri.AbsoluteUri));
        }
    }
}
