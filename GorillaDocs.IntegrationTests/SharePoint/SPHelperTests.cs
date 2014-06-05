using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using GorillaDocs.SharePoint;
using System.Threading;
using System.IO;

namespace GorillaDocs.IntegrationTests
{
    [TestFixture]
    public class SPHelperTests
    {
        const string webUrl = "https://portal.macroview.com.au";
        const string Wrong_Url = "http://somesite.com";
        const string Wrong_Title = "Wrong_Title";
        bool ReturnedToCallback = false;

        [SetUp]
        public void setup() { }

        [Test]
        public void The_Web_should_contain_some_document_libraries()
        {
            var libraries = SPHelper.GetLibraries(webUrl);
            Assert.That(libraries.Count > 0);
        }

        [Test]
        public void The_Web_should_contain_some_document_libraries_returned_asynchronously()
        {
            SPHelper.GetLibraries_Async(webUrl, GetLibrariesSuccessMethod, null);
            WaitForCallback();
        }
        void GetLibrariesSuccessMethod(List<SPLibrary> libraries)
        {
            ReturnedToCallback = true;
            Assert.That(libraries.Count > 0);
        }

        [Test]
        public void An_Exception_is_returned_if_the_webUrl_is_wrong()
        {
            SPHelper.GetLibraries_Async(Wrong_Url, null, GetLibrariesFailureMethod);
            WaitForCallback();
        }
        void GetLibrariesFailureMethod(AggregateException ae)
        {
            ReturnedToCallback = true;
            ae = ae.Flatten();
            foreach (Exception ex in ae.InnerExceptions)
                Assert.That(ex.Message == string.Format("Cannot contact site at the specified URL {0}.", Wrong_Url));
        }

        [Test]
        public void There_should_be_some_documents_in_the_library()
        {
            var files = SPHelper.GetFiles(webUrl, "Precedents");
            Assert.That(files != null);
        }

        [Test]
        public void There_should_be_some_Word_documents_in_the_library()
        {
            var files = SPHelper.GetFiles(webUrl, "Precedents", new String[] { ".doc", ".docx", ".dot", ".dotx" });
            Assert.That(files != null);
        }

        [Test]
        public void That_no_errors_occur_when_downloading_a_file()
        {
            var files = SPHelper.GetFiles(webUrl, "Precedents", new String[] { ".doc", ".docx", ".dot", ".dotx" });
            files.First().Download(new DirectoryInfo(Path.GetTempPath() + "\\GorillaDocs"));
        }

        [Test]
        public void That_the_files_list_serializes()
        {
            var files = SPHelper.GetFiles(webUrl, "Precedents", new String[] { ".doc", ".docx", ".dot", ".dotx" });
            var result = Serializer.SerializeToString(files);
            Assert.That(!string.IsNullOrEmpty(result));
        }

        [Test]
        public void There_should_be_some_documents_in_the_library_returned_asynchronously()
        {
            SPHelper.GetFiles_Async(webUrl, "Precedents", GetFilesSuccessMethod, null);
            WaitForCallback();
        }
        void GetFilesSuccessMethod(List<SPFile> files)
        {
            ReturnedToCallback = true;
            Assert.That(files.Count > 0);
        }

        [Test]
        public void An_Exception_is_returned_if_the_list_Title_is_wrong()
        {
            SPHelper.GetFiles_Async(webUrl, Wrong_Title, null, GetFilesFailureMethod);
            WaitForCallback();
        }
        void GetFilesFailureMethod(AggregateException ae)
        {
            ReturnedToCallback = true;
            ae = ae.Flatten();
            foreach (Exception ex in ae.InnerExceptions)
                Assert.That(ex.Message == string.Format("List '{0}' does not exist at site with URL '{1}'.", Wrong_Title, webUrl));
        }

        void WaitForCallback()
        {
            ReturnedToCallback = false;
            while (true)
            {
                Thread.Sleep(1000);
                if (ReturnedToCallback)
                    return;
            }
        }


    }
}
