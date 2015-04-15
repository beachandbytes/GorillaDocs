using GorillaDocs.Models;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GorillaDocs.SharePoint
{
    public class SPUsers
    {
        public delegate void GetUsersSuccessCallback(List<Contact> users);
        public delegate void GetUserNamesAndIDsSuccessCallback(Dictionary<int, string> users);
        public delegate void GetUserSuccessCallback(Contact user);
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
                    var siteUsers = web.SiteUsers;
                    context.ExecuteQuery();

                    foreach (var user in siteUsers)
                        if (user.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.User)
                            if (user.Title.ToLower().Contains(Filter.ToLower()) && !users.Any(x => x.FullName == user.Title))
                                users.Add(new Contact() { ExternalId = user.Id.ToString(), FullName = user.Title, EmailAddress = user.Email });
                }
            }
            return users;
        }

        public static List<Contact> GetUsersWithProperties(Uri requestUri, string Filter = "")
        {
            ClientContext context;
            var users = new List<Contact>();
            if (ClientContextUtilities.TryResolveClientContext(requestUri, out context, null))
            {
                var userProfilesResult = new List<PersonProperties>();
                using (context)
                {
                    var web = context.Web;
                    var peopleManager = new PeopleManager(context);

                    var siteUsers = from user in web.SiteUsers
                                    where user.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.User
                                    select user;
                    var usersResult = context.LoadQuery(siteUsers);
                    context.ExecuteQuery();

                    foreach (var user in usersResult)
                    {
                        if (user.Title.ToLower().Contains(Filter.ToLower()) && !users.Any(x => x.FullName == user.Title))
                        {
                            var userProfile = peopleManager.GetPropertiesFor(user.LoginName);
                            context.Load(userProfile);
                            userProfilesResult.Add(userProfile);
                        }
                    }
                    context.ExecuteQuery();

                    foreach (var user in usersResult)
                    {
                        var contact = new Contact() { ExternalId = user.Id.ToString(), FullName = user.Title, EmailAddress = user.Email };
                        var userProfile = userProfilesResult.FirstOrDefault(x => x.IsPropertyAvailable("DisplayName") && x.DisplayName == user.Title);
                        if (userProfile != null)
                        {
                            contact.Position = userProfile.IsPropertyAvailable("Title") ? userProfile.Title : string.Empty;
                            contact.PhoneNumber = userProfile.IsPropertyAvailable("UserProfileProperties") && userProfile.UserProfileProperties.ContainsKey("WorkPhone") ? userProfile.UserProfileProperties["WorkPhone"] : string.Empty;
                        }
                        users.Add(contact);
                    }
                }
            }
            return users;
        }

        public static Dictionary<int, string> GetUserNamesAndIds(Uri requestUri, string Filter = "")
        {
            ClientContext context;
            var users = new Dictionary<int, string>();
            if (ClientContextUtilities.TryResolveClientContext(requestUri, out context, null))
            {
                using (context)
                {
                    var web = context.Web;
                    context.Load(web, w => w.SiteUsers);
                    var siteUsers = web.SiteUsers;
                    context.ExecuteQuery();

                    foreach (var user in siteUsers)
                        if (user.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.User)
                            if (user.Title.ToLower().Contains(Filter.ToLower()) && !users.ContainsValue(user.Title))
                                users.Add(user.Id, user.Title);
                }
            }
            return users;
        }

        public static Contact GetUser(Uri requestUri, int Id)
        {
            ClientContext context;
            if (ClientContextUtilities.TryResolveClientContext(requestUri, out context, null))
            {
                using (context)
                {
                    var web = context.Web;
                    context.Load(web);
                    var user = web.GetUserById(Id);
                    context.Load(user, u => u.LoginName);
                    context.ExecuteQuery();

                    var peopleManager = new PeopleManager(context);

                    var userProfile = peopleManager.GetPropertiesFor(user.LoginName);
                    context.Load(userProfile);
                    context.ExecuteQuery();

                    var contact = new Contact() { ExternalId = user.Id.ToString(), FullName = user.Title, EmailAddress = user.Email };
                    if (userProfile.IsPropertyAvailable("Title"))
                        contact.Position = userProfile.Title;
                    if (userProfile.IsPropertyAvailable("UserProfileProperties") && userProfile.UserProfileProperties.ContainsKey("WorkPhone"))
                        contact.PhoneNumber = userProfile.UserProfileProperties["WorkPhone"];
                    return contact;
                }
            }
            throw new InvalidOperationException(string.Format("Unable to find user '{0}' at '{1}'", Id, requestUri.AbsoluteUri));
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

        public static void GetUserNamesAndIds_Async(Uri requestUri, GetUserNamesAndIDsSuccessCallback SuccessCallback, FailureCallback FailureCallback, string Filter = "")
        {
            Task<Dictionary<int, string>> T = Task.Factory.StartNew(() =>
            {
                return GetUserNamesAndIds(requestUri, Filter);
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

        public static void GetUser_Async(Uri requestUri, GetUserSuccessCallback SuccessCallback, FailureCallback FailureCallback, int Id)
        {
            Task<Contact> T = Task.Factory.StartNew(() =>
            {
                return GetUser(requestUri, Id);
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
