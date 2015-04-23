using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs.Tests
{
    [TestFixture]
    public class AutoMapperTests
    {
        public class Person
        {
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public int? Age { get; set; }
            public string Address { get; set; }
        }

        Person source = null;

        [SetUp]
        public void setup()
        {
            source = new Person
            {
                FirstName = "Bill",
                LastName = "Smith",
                Age = 43,
                Address = "123 Some Street"
            };
        }

        [Test]
        public void Test_Merging_Two_Objects()
        {
            var destination = new Person
            {
                FirstName = "Barbara",
                LastName = null,
                Age = 41,
                Address = null
            };

            ObjectMapper.MergeNulls(source,destination);
            Assert.That(destination, Is.Not.Null);
        }

        [Test]
        public void Test_Copying_an_object()
        {
            var destination = new Person();
            ObjectMapper.Copy(source, destination);
            Assert.That(destination, Is.Not.Null);
        }

    }
}
