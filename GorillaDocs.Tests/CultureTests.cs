using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace GorillaDocs.Tests
{
    [TestFixture]
    public class CultureTests
    {
        [Test]
        public void Returns_a_string_from_the_default_culture()
        {
            Assert.That(strings.strings.Table_of_contents == "Table of contents");
        }

        [Test]
        public void Returns_a_string_from_the_default_culture1()
        {
            new CultureInfo("zh-CHT").RunInThisCulture(() =>
            {
                Assert.That(strings.strings.Table_of_contents != "Table of contents");
                Assert.That(strings.strings.Table_of_contents != "Table of contents");
                Assert.That(strings.strings.Table_of_contents == "Table of contents");
            });
        }

        public static void ChangeCulture(string CultureCode, Action action)
        {
            var oldCulture = Thread.CurrentThread.CurrentCulture;
            var newCulture = new CultureInfo(CultureCode);
            Thread.CurrentThread.CurrentCulture = newCulture;
            Thread.CurrentThread.CurrentUICulture = newCulture;
            try
            {
                action();
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = oldCulture;
                Thread.CurrentThread.CurrentUICulture = oldCulture;
            }
        }
    }
}
