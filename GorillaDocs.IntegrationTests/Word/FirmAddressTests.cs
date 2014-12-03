using GorillaDocs.IntegrationTests.Helpers;
using GorillaDocs.Word;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.IntegrationTests.Word
{
    [TestFixture]
    public class FirmAddressTests
    {
        Wd.Application wordApp;
        Wd.Document doc;

        [SetUp]
        public void setup()
        {
            wordApp = WordApplicationHelper.GetApplication();
            doc = wordApp.Documents.Open(IOHelpers.ContentControlsData.FullName);
        }

        [Test]
        public void Test_inserting_Firm_address()
        {
            
        }
    }
}
