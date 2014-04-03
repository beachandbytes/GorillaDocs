using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using System.Reflection;

namespace GorillaDocs.Tests
{
    [TestFixture]
    public class AssemblyHelperTests
    {
        const string expected_Title = "GorillaDocs.Tests";
        const string expected_FileVersion = "1.0.0.0";
        const string expected_PathEnding = @"\GorillaDocs\GorillaDocs.Tests\bin\debug";
        Assembly assembly;

        [SetUp]
        public void setup()
        {
            assembly = Assembly.GetExecutingAssembly();
        }

        [Test]
        public void Title_returns_GorillaDocsTests()
        {
            Assert.That(assembly.Title() == expected_Title);
        }

        [Test]
        public void FileVersion_returns_1000()
        {
            Assert.That(assembly.FileVersion() == expected_FileVersion);
        }

        [Test]
        public void Path_ends_with_GorrilaTests()
        {
            Assert.That(assembly.Path().EndsWith(expected_PathEnding, StringComparison.OrdinalIgnoreCase));
        }
    }
}
