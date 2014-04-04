using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using GorillaDocs;
using System.Xml.Linq;
using System.Xml;

namespace GorillaDocs.Tests
{
    [TestFixture]
    public class SerializerTests
    {
        Person person = new Person() { FirstName = "John", LastName = "Smith" };
        string personXml = "<?xml version=\"1.0\" encoding=\"utf-16\"?>\r\n<Person>\r\n  <FirstName>John</FirstName>\r\n  <LastName>Smith</LastName>\r\n</Person>";

        [SetUp]
        public void setup() { }

        [Test]
        public void The_class_is_Serialized_in_the_correct_format()
        {
            var result = Serializer.SerializeToString<Person>(person);
            Assert.AreEqual(personXml, result);
        }

        [Test]
        public void The_result_does_not_contain_unnecessary_namespaces()
        {
            var result = Serializer.SerializeToString<Person>(person);
            Assert.That(!result.Contains("xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\""));
            Assert.That(!result.Contains("xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\""));
        }
    }

    public class Person
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
    }
}
