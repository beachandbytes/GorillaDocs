using GorillaDocs.Models;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GorillaDocs.SharePoint
{
    public class SPUsers
    {
        public delegate void GetUsersSuccessCallback(List<Contact> users);
        public delegate void FailureCallback(AggregateException ae);

        public static List<Contact> GetUsers(Uri requestUri, string Filter = "")
        {
            ClientContext context;
            var users = new List<Contact>();
            if (ClientContextUtilities.TryResolveClientContext(requestUri, out context, null))
            {
                using (context)
                {
                    var web = context.Web;
                    context.Load(web, w => w.Title, w => w.Description, w => w.SiteUsers);
                    //IEnumerable<User> siteUsers;
                    ////UserCollection siteUsers;
                    //if (string.IsNullOrEmpty(Filter))
                    var siteUsers = web.SiteUsers;
                    //else
                    //{
                    //    //var query = from user in web.SiteUsers.Include(u => u.Title, u => u.Email)
                    //    //            where user.Title.ToLower().IndexOf(Filter.ToLower()) >= 0
                    //    //            select user;
                    //    //var query = web.SiteUsers.Include(u => u.Title, u => u.Email).Where(u => u.Title.ToLower().IndexOf(Filter.ToLower()) >= 0);
                    //    var query = web.SiteUsers.Include(u => u.Title, u => u.Email).Where(u => u.Title.ToLower().StartsWith(Filter.ToLower()));
                    //    siteUsers = context.LoadQuery(query);
                    //}
                    context.ExecuteQuery();

                    foreach (var user in siteUsers)
                        if (user.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.User)
                            if (user.Title.ToLower().Contains(Filter.ToLower()) && !users.Any(x => x.FullName == user.Title))
                                users.Add(new Contact()
                                {
                                    FullName = user.Title,
                                    EmailAddress = user.Email
                                });
                }
            }
            return users;
        }

        public static void GetUsers_Async(Uri requestUri, GetUsersSuccessCallback SuccessCallback, FailureCallback FailureCallback, string Filter = "")
        {
            Task<List<Contact>> T = Task.Factory.StartNew(() =>
            {
                return GetUsers(requestUri, Filter);
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
