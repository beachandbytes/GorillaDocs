using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using GorillaDocs.SharePoint;
using System.Threading;
using System.IO;
using System.Net;
using System.Text;
using System.Runtime.InteropServices;
using System.Security;

namespace GorillaDocs.IntegrationTests
{
    [TestFixture]
    public class SPHelperTests
    {
        const string webUrl = "http://mvsp13.macroview.com.au";
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
            //AuthPrompt prompt = new AuthPrompt(webUrl);
            //if (prompt.PromptForPassword())
            //    Assert.That(prompt.User != null);

            var files = SPHelper.GetFiles(webUrl, "Precedents"); //, null, new NetworkCredential("", ""));
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
        public void That_no_errors_occur_when_downloading_a_file_with_spaces_in_the_name()
        {
            var file = new SPFile()
            {
                RemoteUrl = "http://mvsp13.macroview.com.au/sites/Products/Manuals/Alison Campbell_23Apr12 13.16.45_DMF issue log.msg",
                Name = "Alison Campbell_23Apr12 13.16.45_DMF issue log",
                Extension = ".msg"
            };
            file.Download(new DirectoryInfo("C:\\Users\\Matthew\\AppData\\Local\\Temp\\MacroView.Office\\3486fefb-85f7-4bd6-bc37-1b418bd0d7e4"));
        }

        [Test]
        public void That_the_files_list_serializes()
        {
            var files = SPHelper.GetFiles(webUrl, "Precedents", new String[] { ".doc", ".docx", ".dot", ".dotx" });
            var result = Serializer.SerializeToString(files);
            Assert.That(!string.IsNullOrEmpty(result));
        }

        [Test]
        public void That_a_library_without_Category_works()
        {
            var files = SPHelper.GetFiles(webUrl + "/sites/products", "Manuals", new String[] { ".doc", ".docx", ".dot", ".dotx" });
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

        [Test]
        public void Get_the_Term_Stores()
        {
            var termStores = TaxonomyHelper.GetTermStores("http://mvuatsp13.macroview.com.au/");
            Assert.That(termStores != null);
        }

        [Test]
        public void Get_the_Term_Groups()
        {
            var termGroups = TaxonomyHelper.GetTermGroups("http://mvuatsp13.macroview.com.au/", new Guid("2a09029f-199e-4013-9289-e4b00fb0ea14"));
            Assert.That(termGroups != null);
        }

        [Test]
        public void Get_the_Term_Sets()
        {
            var termSets = TaxonomyHelper.GetTermSets("http://mvuatsp13.macroview.com.au/", new Guid("2a09029f-199e-4013-9289-e4b00fb0ea14"), new Guid("edde64cc-6116-4e86-b509-2b9a05aeb57d"));
            Assert.That(termSets != null);
        }

        [Test]
        public void Get_term_id()
        {
            var id = TaxonomyHelper.GetTermId(new Uri("http://tests2012sp2013.macroview.com.au/sites/caseandmatter/Atlas%20Funds%20Management"), "Letter");
        }

        [Test]
        public void Get_Users()
        {
            var users = SPUsers.GetUsersWithProperties(new Uri("https://portal.macroview.com.au/"));
            //var users = SPUsers.GetUsers(new Uri("http://mvuatsp13.macroview.com.au/"));
            //var users = SPUsers.GetUsersWithProperties(new Uri("http://tests2012sp2013.macroview.com.au/sites/caseandmatter/Atlas%20Funds%20Management"));
            //var users = SPUsers.GetUsers("http://tests2012sp2013.macroview.com.au/sites/caseandmatter/");
            Assert.That(users != null);
        }

        [Test]
        public void Get_filtered_Users()
        {
            var users = SPUsers.GetUsers(new Uri("http://mvuatsp13.macroview.com.au/"), "Matt");
            Assert.That(users != null);
        }
    }

    public class AuthPrompt
    {
        public struct CREDUI_INFO
        {
            public int cbSize;
            public IntPtr hwndParent;
            public string pszMessageText;
            public string pszCaptionText;
            public IntPtr hbmBanner;
        }

        [DllImport("credui")]
        private static extern CredUIReturnCodes CredUIPromptForCredentials(ref CREDUI_INFO creditUR,
              string targetName,
              IntPtr reserved1,
              int iError,
              StringBuilder userName,
              int maxUserName,
              StringBuilder password,
              int maxPassword,
              [MarshalAs(UnmanagedType.Bool)] ref bool pfSave,
              CREDUI_FLAGS flags);

        [Flags]
        enum CREDUI_FLAGS
        {
            INCORRECT_PASSWORD = 0x1,
            DO_NOT_PERSIST = 0x2,
            REQUEST_ADMINISTRATOR = 0x4,
            EXCLUDE_CERTIFICATES = 0x8,
            REQUIRE_CERTIFICATE = 0x10,
            SHOW_SAVE_CHECK_BOX = 0x40,
            ALWAYS_SHOW_UI = 0x80,
            REQUIRE_SMARTCARD = 0x100,
            PASSWORD_ONLY_OK = 0x200,
            VALIDATE_USERNAME = 0x400,
            COMPLETE_USERNAME = 0x800,
            PERSIST = 0x1000,
            SERVER_CREDENTIAL = 0x4000,
            EXPECT_CONFIRMATION = 0x20000,
            GENERIC_CREDENTIALS = 0x40000,
            USERNAME_TARGET_CREDENTIALS = 0x80000,
            KEEP_USERNAME = 0x100000,
        }

        public enum CredUIReturnCodes
        {
            NO_ERROR = 0,
            ERROR_CANCELLED = 1223,
            ERROR_NO_SUCH_LOGON_SESSION = 1312,
            ERROR_NOT_FOUND = 1168,
            ERROR_INVALID_ACCOUNT_NAME = 1315,
            ERROR_INSUFFICIENT_BUFFER = 122,
            ERROR_INVALID_PARAMETER = 87,
            ERROR_INVALID_FLAGS = 1004,
        }

        /// <summary>
        /// Prompts for password.
        /// </summary>
        /// <param name="user">The user.</param>
        /// <param name="password">The password.</param>
        /// <returns>True if no errors.</returns>
        public bool PromptForPassword()
        {
            // Setup the flags and variables
            StringBuilder userPassword = new StringBuilder(), userID = new StringBuilder();
            CREDUI_INFO credUI = new CREDUI_INFO();
            credUI.cbSize = Marshal.SizeOf(credUI);
            bool save = true;
            //bool save = true;
            //CREDUI_FLAGS flags = CREDUI_FLAGS.ALWAYS_SHOW_UI | CREDUI_FLAGS.GENERIC_CREDENTIALS | CREDUI_FLAGS.DO_NOT_PERSIST | CREDUI_FLAGS.SHOW_SAVE_CHECK_BOX | CREDUI_FLAGS.PERSIST;
            CREDUI_FLAGS flags = CREDUI_FLAGS.ALWAYS_SHOW_UI | CREDUI_FLAGS.GENERIC_CREDENTIALS | CREDUI_FLAGS.SHOW_SAVE_CHECK_BOX;

            // Prompt the user
            CredUIReturnCodes returnCode = CredUIPromptForCredentials(ref credUI, this.serverName, IntPtr.Zero, 0, userID, 100, userPassword, 100, ref save, flags);

            _user = userID.ToString();

            _securePassword = new SecureString();
            for (int i = 0; i < userPassword.Length; i++)
            {
                _securePassword.AppendChar(userPassword[i]);
            }

            return (returnCode == CredUIReturnCodes.NO_ERROR);
        }

        private string _user;
        public string User
        {
            get { return _user; }
        }

        private SecureString _securePassword;
        public SecureString SecurePassword
        {
            get { return _securePassword; }
        }
        private string serverName;

        public AuthPrompt(string serverName)
        {
            this.serverName = serverName;
        }
    }
}
