using GorillaDocs.Models;
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
            public Address Address { get; set; }
            public Contact Friend { get; set; }
        }

        public class Address
        {
            public string Street { get; set; }
            public string City { get; set; }
            public string State { get; set; }
            public int PostCode { get; set; }
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
                Address = new Address() { Street = "123 Some Street", City = "OurTown", State = "State", PostCode = 1000 },
                Friend = new Contact() { FullName = "Someone", StreetAddress1 = "SA1", PhoneNumber = "234" }
            };
        }

        [Test]
        public void Test_Merging_Two_Objects()
        {
            var destination = new Person
            {
                FirstName = "Barbara",
                LastName = null,
                Age = null,
                Address = null,
                Friend = new Contact()
            };

            ObjectMapper.MergeNulls(source, destination);
            ObjectMapper.MergeNulls(source.Friend, destination.Friend);
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
