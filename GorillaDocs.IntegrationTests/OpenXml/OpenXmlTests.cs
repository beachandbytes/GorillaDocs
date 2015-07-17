using GorillaDocs.OpenXml;
using NUnit.Framework;
using System.IO;
using System.Reflection;

namespace GorillaDocs.IntegrationTests.OpenXml
{
    [TestFixture]
    public class OpenXmlTests
    {
        [Test]
        public void test1()
        {
            var doc = new WordDocument(SampleFile);
            var variables = doc.Variables;
        }

        static FileInfo SampleFile { get { return new FileInfo(Assembly.GetExecutingAssembly().Path() + @"\OpenXml\SampleData\Agreement Front Cover.docx"); } }
    }
}
