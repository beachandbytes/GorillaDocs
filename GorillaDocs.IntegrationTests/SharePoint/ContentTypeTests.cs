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
            const string webUrl = "http://mvuatsp13.macroview.com.au/sites/mvcm1/CMLibrary/Forms/All%20Matters.aspx";
            const string listTitle = "CMLibrary - CM Library";
            var result = ContentTypes.GetContentTypes(new Uri(webUrl), listTitle);
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
