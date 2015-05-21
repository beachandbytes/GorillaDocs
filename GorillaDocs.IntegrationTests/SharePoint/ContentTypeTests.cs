using GorillaDocs.SharePoint;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs.IntegrationTests.SharePoint
{
    [TestFixture]
    public class ContentTypeTests
    {
        [Test]
        public void Get_Content_Types_on_portal()
        {
            const string webUrl = "https://portal.macroview.com.au/Documents";
            const string listTitle = "Documents";
            var result = ContentTypes.GetContentTypes(new Uri(webUrl), listTitle);
            Assert.That(result != null);
        }

        [Test]
        public void Get_Content_Types_on_mvuatsp13()
        {
            //const string webUrl = "http://mvuatsp13.macroview.com.au/sites/mvcm1/";
            const string webUrl = "http://mvuatsp2010.macroview.com.au/";
            //const string listTitle = "CMLibrary - CM Library";
            const string listTitle = "Documents";
            var credentials = new System.Net.NetworkCredential("MacroView\\mjf", "timnmvid&");
            var result = ContentTypes.GetContentTypes(new Uri(webUrl), listTitle, credentials);
            Assert.That(result != null);
        }

        [Test]
        public void Get_Users_mvuatsp13()
        {
            const string webUrl = "http://mvuatsp13.macroview.com.au/sites/mvcm1/";
            const string listTitle = "CMLibrary - CM Library";
            var credentials = new System.Net.NetworkCredential("MacroView\\mjf", "timnmvid&");
            var result = SPUsers.GetUsers(new Uri(webUrl), "", credentials);
            Assert.That(result != null);
        }

        [Test]
        public void Get_Users_mvuatsp2010()
        {
            const string webUrl = "http://mvuatsp2010.macroview.com.au/";
            const string listTitle = "Documents";
            var credentials = new System.Net.NetworkCredential("MacroView\\mjf", "timnmvid&");
            var result = SPUsers.GetUsers(new Uri(webUrl), "", credentials);
            Assert.That(result != null);
        }

        [Test]
        public void Get_Taxonomy_mvuatsp13()
        {
            const string webUrl = "http://mvuatsp13.macroview.com.au/sites/mvcm1/";
            var credentials = new System.Net.NetworkCredential("MacroView\\mjf", "timnmvid&");
            var result = TaxonomyHelper.GetTermStores(webUrl, credentials);
            Assert.That(result != null);
        }

        [Test]
        public void Get_Taxonomy_mvuatsp2010()
        {
            const string webUrl = "http://mvuatsp2010.macroview.com.au/";
            var credentials = new System.Net.NetworkCredential("MacroView\\mjf", "timnmvid&");
            var result = TaxonomyHelper.GetTermStores(webUrl, credentials);
            Assert.That(result != null);
        }

        [Test]
        public void Does_the_content_type_exist_on_portal()
        {
            const string webUrl = "https://portal.macroview.com.au/Documents";
            const string listTitle = "Documents";
            const string contentTypeName = "test";
            var result = ContentTypes.ContentTypeExists(new Uri(webUrl), listTitle, contentTypeName);
            Assert.That(result);
        }

        [Test]
        public void Does_the_content_type_exist_on_mvuatsp13()
        {
            const string webUrl = "http://mvuatsp13.macroview.com.au/sites/mvcm1/CMLibrary/Forms/All%20Matters.aspx";
            const string listTitle = "CMLibrary - CM Library";
            const string contentTypeName = "Matter Document Set";
            var result = ContentTypes.ContentTypeExists(new Uri(webUrl), listTitle, contentTypeName);
            Assert.That(result);
        }
    }
}
