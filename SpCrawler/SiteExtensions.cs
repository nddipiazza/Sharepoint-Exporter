using System;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using System.Linq.Expressions;

namespace SpPrefetchIndexBuilder
{
    public static class SiteExtensions
    {
        public static List<Web> EnumAllWebs(this Site site, params Expression<Func<Web, object>>[] retrievals)
        {
            var ctx = site.Context;
            var rootWeb = site.RootWeb;
            ctx.Load(rootWeb, retrievals);
            var result = new List<Web>();
            result.Add(rootWeb);
            EnumAllWebsInner(rootWeb, result, retrievals);
            return result;
        }

        private static void EnumAllWebsInner(Web parentWeb, ICollection<Web> result, params Expression<Func<Web, object>>[] retrievals)
        {
            try
            {
                var ctx = parentWeb.Context;
                var webs = parentWeb.Webs;
                ctx.Load(webs, wcol => wcol.Include(retrievals));
                ctx.ExecuteQuery();
                foreach (var web in webs)
                {
                    result.Add(web);
                    EnumAllWebsInner(web, result, retrievals);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Could not fetch site due to error {0}", ex);
            }
        }
    }
}
