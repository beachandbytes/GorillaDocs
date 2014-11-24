using GorillaDocs.Word;
using NUnit.Framework;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.IntegrationTests
{
    [TestFixture]
    public class CloneTests
    {
        //const string template = @"C:\Repos\MacroView\VS2010\BMGlobal.Office\BMGlobal.Common\Office Folders\Templates\Global\Correspondence\Letter.dotm";
        const string template = @"C:\Repos\MacroView\VS2010\BMGlobal.Office\BMGlobal.Common\Office Folders\Templates\Global\Business Development\Marketing Pitch.dotm";
        //const string template = @"C:\Repos\MacroView\VS2010\BMGlobal.Office\BMGlobal.Common\Office Folders\Templates\Global\Agreements\Global Agreement.dotm";
        Wd.Application wordApp;

        [SetUp]
        public void setup()
        {
            wordApp = WordApplicationHelper.GetApplication();
        }

        [Test]
        public void test1()
        {
            var doc = wordApp.Documents.Add();
            doc.CloneFrom(template);
        }
    }
}
