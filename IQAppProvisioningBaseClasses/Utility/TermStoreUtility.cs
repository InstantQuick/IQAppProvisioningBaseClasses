using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace IQAppProvisioningBaseClasses.Utility
{
    public class TermStoreUtility
    {
        public static TermStore GetTermStore(ClientContext ctx, TaxonomySession ts)
        {
            //It has been observed that GetDefault doesn't work in some service proxy configurations
            //so this attempts to fall back to the one at index 0 if GetDefaultSiteCollectionTermStore fails
            var termStore = ts.GetDefaultSiteCollectionTermStore();
            ctx.Load(termStore, t => t.Id);
            ctx.Load(ts.TermStores);
            ctx.ExecuteQueryRetry();

            if (!termStore.IsPropertyAvailable("Id") && ts.TermStores.Count > 0)
            {
                termStore = ts.TermStores[0];
            }
            return termStore;
        }
    }
}

