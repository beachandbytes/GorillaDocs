using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;

namespace GorillaDocs.Tests
{
    [TestFixture]
    public class InternationalAddressFormatsTests
    {
        readonly InternationalAddressFormats address = new InternationalAddressFormats("123 Some Street", "Sydney", "NSW", "2000", "Australia");
        const string expected_fr_FR = "123 Some Street\n2000 Sydney NSW\nAustralia";
        const string expected_es_ES = "123 Some Street\n2000 Sydney\nAustralia";
        const string expected_ro_RO = "123 Some Street\n2000 Sydney\nNSW\nAustralia";
        const string expected_ru_RU = "Australia\n2000\nNSW Sydney\n123 Some Street";
        const string expected_en_AU = "123 Some Street\nSydney NSW 2000\nAustralia";
        const string expected_en_US = "123 Some Street\nSydney, NSW 2000\nAustralia";
        
        [SetUp]
        public void setup() { }

        [Test]
        public void GetAddress_returns_the_en_AU_format_when_no_culture_provided()
        {
            Assert.That(address.GetAddress() == expected_en_AU);
        }

        [Test]
        public void GetAddress_returns_the_fr_FR_format()
        {
            Assert.That(address.GetAddress("fr-FR") == expected_fr_FR);
        }

        [Test]
        public void GetAddress_returns_the_es_ES_format()
        {
            Assert.That(address.GetAddress("es-ES") == expected_es_ES);
        }

        [Test]
        public void GetAddress_returns_the_ro_RO_format()
        {
            Assert.That(address.GetAddress("ro-RO") == expected_ro_RO);
        }

        [Test]
        public void GetAddress_returns_the_ru_RU_format()
        {
            Assert.That(address.GetAddress("ru-RU") == expected_ru_RU);
        }

        [Test]
        public void GetAddress_returns_the_en_AU_format()
        {
            Assert.That(address.GetAddress("en-AU") == expected_en_AU);
        }

        [Test]
        public void GetAddress_returns_the_en_US_format()
        {
            Assert.That(address.GetAddress("en-US") == expected_en_US);
        }
    }
}
