using GorillaDocs.SharePoint;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs.IntegrationTests.SharePoint
{
    [TestFixture]
    public class CheckInTests
    {
        [Test]
        public void Discard_CheckOut()
        {
            CheckInHelper.DiscardCheckOut("http://mvuatsp2010.macroview.com.au", "http://mvuatsp2010.macroview.com.au/MacroView Library Test/MacroView Web Site.doc");
        }
    }
}
