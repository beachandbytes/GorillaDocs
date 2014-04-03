using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;

namespace GorillaDocs.Tests
{
    [TestFixture]
    public class CollectionHelperTests
    {
        List<string> sourceList = new List<string>() { "Test", "One", "Two", "Test", "Three" };
        List<int> sourceInts = new List<int>() { 1, 3, 2, 6, 7 };
        List<int> emptyList = new List<int>();

        [SetUp]
        public void setup()
        {
        }

        [Test]
        public void Replace_X_with_Y_and_return_Y()
        {
            Assert.That(sourceList.ReplaceAndReturn("One", "Four") == "Four");
        }

        [Test]
        public void RemoveAll_removes_2()
        {
            Assert.That(sourceList.RemoveAll<string>(x => x.Equals("Test")) == 2);
        }

        [Test]
        public void Return_the_first_item_in_the_collection()
        {
            Assert.That(sourceInts.FirstOrCreateIfEmpty() == 1); 
        }

        [Test]
        public void Return_a_new_item()
        {
            Assert.That(emptyList.FirstOrCreateIfEmpty() == 0);
        }
    }
}
