using GorillaDocs.IntegrationTests.Helpers;
using GorillaDocs.Word;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.IntegrationTests
{
    [TestFixture]
    public class ContentControlTests
    {
        Wd.Application wordApp;
        Wd.Document doc;

        [SetUp]
        public void setup()
        {
            this.wordApp = WordApplicationHelper.GetWordApplication();
            this.doc = this.wordApp.Documents.Open(IOHelpers.ContentControlsData.FullName);
        }

        [Test]
        public void Move_the_range_out_of_the_first_content_control()
        {
            Wd.Range range = this.doc.ContentControls[1].Range;
            range.MoveOutOfContentControl();
            Assert.That(range.ContentControls.Count == 0); 
        }

        [Test]
        public void Move_the_range_out_of_a_content_control()
        {
            Wd.Range range = this.doc.Range();
            range.Collapse(Wd.WdCollapseDirection.wdCollapseStart);
            range.Move(Wd.WdUnits.wdCharacter, 3);
            range.MoveOutOfContentControl();
            Assert.That(range.ContentControls.Count == 0);
        }

        [Test]
        public void Select_multiple_controls_and_move_the_range_out_of_the_content_control()
        {
            Wd.Range range = this.doc.Range();
            range.Start = this.doc.ContentControls[1].Range.Start;
            range.End = this.doc.ContentControls[this.doc.ContentControls.Count].Range.End;
            range.MoveOutOfContentControl();
            Assert.That(range.ContentControls.Count == 0);
        }

        [Test]
        public void Move_out_of_controls_in_a_table_cell()
        {
            Wd.Range range = this.doc.ContentControls.Last().Range;
            range.MoveOutOfContentControl();
            range.InsertParagraphAfter();
            Assert.That(range.ContentControls.Count == 0);
        }

    }
}
