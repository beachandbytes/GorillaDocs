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
            if (wordApp.Documents.Exists(IOHelpers.FirmAddressTestData))
                wordApp.Documents.First(IOHelpers.FirmAddressTestData).Close(Wd.WdSaveOptions.wdDoNotSaveChanges);
            doc = wordApp.Documents.Open(IOHelpers.FirmAddressTestData.FullName);
        }

        [Test]
        public void Update_all_Firm_addresses_in_the_body()
        {
            doc.Range().UpdateFirmAddressesControls(IOHelpers.FirmAddressData);
        }

        [Test]
        public void Update_all_firm_addresses_in_the_headers()
        {
            foreach (Wd.Section section in doc.Sections)
                section.Headers.UpdateFirmAddressesControls(IOHelpers.FirmAddressData);
        }

        [Test]
        public void Update_all_firm_addresses_in_the_footers()
        {
            foreach (Wd.Section section in doc.Sections)
                section.Footers.UpdateFirmAddressesControls(IOHelpers.FirmAddressData);
        }

        [Test]
        public void Update_all_firm_addresses_in_the_shapes()
        {
            doc.Shapes.UpdateFirmAddressesControls(IOHelpers.FirmAddressData);
        }

        [Test]
        public void Update_all_firm_addresses_in_the_shapes_in_the_headers()
        {
            doc.Sections[1].Headers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.UpdateFirmAddressesControls(IOHelpers.FirmAddressData);
        }
    }
}
