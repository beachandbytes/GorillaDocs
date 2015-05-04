using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GorillaDocs.SharePoint
{
    public class ContentTypes
    {
        public delegate void GetContentTypesSuccessCallback(List<SPContentType> users);
        public delegate void ContentTypeExistsSuccessCallback(bool result);
        public delegate void FailureCallback(AggregateException ae);

        public static List<SPContentType> GetContentTypes(Uri requestUri, string ListTitle)
        {
            ClientContext context;
            var contentTypes = new List<SPContentType>();
            if (ClientContextUtilities.TryResolveClientContext(requestUri, out context, null))
            {
                using (context)
                {
                    var list = context.Web.Lists.GetByTitle(ListTitle);

                    var contentTypeCol = list.ContentTypes;
                    context.Load(contentTypeCol, lstc => lstc.Include(lc => lc.Name));
                    context.ExecuteQuery();

                    foreach (var contentType in contentTypeCol)
                        contentTypes.Add(new SPContentType() { Name = contentType.Name });
                }
            }
            return contentTypes;
        }

        public static bool ContentTypeExists(Uri requestUri, string ListTitle, string ListContentTypeField)
        {
            ClientContext context;
            var contentTypes = new List<SPContentType>();
            if (ClientContextUtilities.TryResolveClientContext(requestUri, out context, null))
            {
                using (context)
                {
                    var list = context.Web.Lists.GetByTitle(ListTitle);

                    var contentTypeCol = list.ContentTypes;
                    context.Load(contentTypeCol, lstc => lstc.Include(lc => lc.Name).Where(lc => lc.Name == ListContentTypeField));
                    context.ExecuteQuery();
                    return contentTypeCol.Count == 1;
                }
            }
            return false;
        }

        public static void GetContentTypes_Async(Uri requestUri, string ListTitle, GetContentTypesSuccessCallback SuccessCallback, FailureCallback FailureCallback)
        {
            Task<List<SPContentType>> T = Task.Factory.StartNew(() =>
            {
                return GetContentTypes(requestUri, ListTitle);
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

        public static void ContentTypeExists_Async(Uri requestUri, string ListTitle, string ListContentTypeField, ContentTypeExistsSuccessCallback SuccessCallback, FailureCallback FailureCallback)
        {
            Task<bool> T = Task.Factory.StartNew(() =>
            {
                return ContentTypeExists(requestUri, ListTitle, ListContentTypeField);
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

        //public dynamic isExist_Helper(ClientContext context, String fieldToCheck, String type)
        //{
        //    var isExist = 0;
        //    ListCollection listCollection = context.Web.Lists;
        //    ContentTypeCollection cntCollection = context.Web.ContentTypes;
        //    FieldCollection fldCollection = context.Web.Fields;
        //    switch (type)
        //    {
        //        case "list":
        //            context.Load(listCollection, lsts => lsts.Include(list => list.Title).Where(list => list.Title == fieldToCheck));
        //            context.ExecuteQuery();
        //            isExist = listCollection.Count;
        //            break;
        //        case "contenttype":
        //            context.Load(cntCollection, cntyp => cntyp.Include(ct => ct.Name).Where(ct => ct.Name == fieldToCheck));
        //            context.ExecuteQuery();
        //            isExist = cntCollection.Count;
        //            break;
        //        case "contenttypeName":
        //            context.Load(cntCollection, cntyp => cntyp.Include(ct => ct.Name, ct => ct.Id).Where(ct => ct.Name == fieldToCheck));
        //            context.ExecuteQuery();
        //            foreach (ContentType ct in cntCollection)
        //            {
        //                return ct.Id.ToString();
        //            }
        //            break;
        //        case "field":
        //            context.Load(fldCollection, fld => fld.Include(ft => ft.Title).Where(ft => ft.Title == fieldToCheck));
        //            try
        //            {
        //                context.ExecuteQuery();
        //                isExist = fldCollection.Count;
        //            }
        //            catch (Exception e)
        //            {
        //                if (e.Message == "Unknown Error")
        //                {
        //                    isExist = fldCollection.Count;
        //                }
        //            }
        //            break;
        //        case "listcntype":
        //            List lst = context.Web.Lists.GetByTitle(fieldToCheck);
        //            ContentTypeCollection lstcntype = lst.ContentTypes;
        //            context.Load(lstcntype, lstc => lstc.Include(lc => lc.Name).Where(lc => lc.Name == fieldToCheck));
        //            context.ExecuteQuery();
        //            isExist = lstcntype.Count;
        //            break;
        //    }
        //    return isExist;
        //}
    }
}
