using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GorillaDocs.Tests
{
    [TestFixture]
    public class DrawingTests
    {
        [Test]
        public void blah()
        {
            var icon = new IconHelper(@"C:\Users\Matthew\Dropbox\W\tfs\All MacroView Projects\MacroView.Office\MacroView.Office\Office Files\Templates\Presentations\MacroView Presentation.potx");
            icon.SaveAsPng(@"C:\Test.png");
        }
    }
}
