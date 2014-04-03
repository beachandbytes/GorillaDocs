using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using GorillaDocs;
using System.IO;
using System.Reflection;

namespace GorillaDocs.Tests
{
    [TestFixture]
    public class IOHelperTests
    {
        DirectoryInfo folder;
        FileInfo file;

        [SetUp]
        public void setup()
        {
            file = new FileInfo(Assembly.GetExecutingAssembly().Path() + @"\GorillaDocs.Test.dll");
            folder = new DirectoryInfo(Assembly.GetExecutingAssembly().Path());
        }

        [Test]
        public void Returns_the_name_of_the_file_without_the_extension()
        {
            Assert.That(file.NameWithoutExtension() == "GorillaDocs.Test");
        }

        [Test]
        public void Returns_the_path_to_the_file()
        {
            Assert.That(file.Path() == Assembly.GetExecutingAssembly().Path());
        }

        [Test]
        public void The_debug_folder_contains_dll_files()
        {
            Assert.That(folder.ContainsFiles("*.dll"));
        }

        [Test]
        public void The_debug_folder_does_not_contain_docx_files()
        {
            Assert.That(!folder.ContainsFiles("*.docx"));
        }
    }
}
