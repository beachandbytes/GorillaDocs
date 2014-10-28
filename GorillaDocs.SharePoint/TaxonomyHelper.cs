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
