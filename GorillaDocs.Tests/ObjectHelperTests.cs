using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using GorillaDocs;

namespace GorillaDocs.Tests
{
    [TestFixture]
    public class ObjectHelperTests
    {
        const string emptyString = "";
        const string nullString = null;

        [Test]
        public void Returns_the_type_of_the_object()
        {
            Assert.That(emptyString.NullableGetType() == typeof(string));
        }

        [Test]
        public void Returns_null_if_the_object_is_null()
        {
            Assert.Null(nullString.NullableGetType());
        }
    }
}
