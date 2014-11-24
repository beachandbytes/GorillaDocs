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
            this.wordApp = WordApplicationHelper.GetApplication();
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

        [Test]
        public void Return_false_if_not_in_a_content_control()
        {
            this.doc.Bookmarks["Not_in_control"].Select();
            Wd.Range range = this.wordApp.Selection.Range;
            Assert.IsFalse(range.InContentControl());
        }

        [Test]
        public void Return_true_if_in_a_content_control()
        {
            Wd.Range range = this.doc.Range().CollapseStart();
            range.Move(Wd.WdUnits.wdCharacter, 3);
            Assert.IsTrue(range.InContentControl());
        }

        [Test]
        public void Test_the_boundaries_of_the_content_control()
        {
            Wd.Range range = this.doc.ContentControls[2].Range;
            Assert.IsTrue(range.InContentControl());

            range.MoveStart(Wd.WdUnits.wdCharacter, -1);
            Assert.IsTrue(range.InContentControlOrContainsControls());

            range.MoveEnd(Wd.WdUnits.wdCharacter, 1);
            Assert.IsTrue(range.InContentControlOrContainsControls());

            range.MoveStart(Wd.WdUnits.wdCharacter, -1);
            Assert.IsTrue(range.InContentControlOrContainsControls());

            range.MoveEnd(Wd.WdUnits.wdCharacter, 1);
            Assert.IsTrue(range.InContentControlOrContainsControls());
        }

        [Test]
        public void Return_true_if_range_includes_many_content_controls()
        {
            Wd.Range range = this.doc.Range();
            Assert.IsTrue(range.InContentControlOrContainsControls());
        }

        [Test]
        public void Return_null_if_not_in_a_content_control()
        {
            this.doc.Bookmarks["Not_in_control"].Select();
            Wd.Range range = this.wordApp.Selection.Range;
            Assert.IsNull(range.GetSurroundingContentControl());
        }

        [Test]
        public void Return_control_if_in_a_content_control()
        {
            Wd.Range range = this.doc.Range().CollapseStart();
            range.Move(Wd.WdUnits.wdCharacter, 3);
            Assert.IsNotNull(range.GetSurroundingContentControl());
        }

        [Test]
        public void Test_the_boundaries_of_the_content_control_not_null()
        {
            Wd.Range range = this.doc.ContentControls[2].Range;
            Assert.IsNotNull(range.GetSurroundingContentControl());

            range.MoveStart(Wd.WdUnits.wdCharacter, -1);
            Assert.IsNotNull(range.GetSurroundingContentControl());

            range.MoveEnd(Wd.WdUnits.wdCharacter, 1);
            Assert.IsNotNull(range.GetSurroundingContentControl());

            range.MoveStart(Wd.WdUnits.wdCharacter, -1);
            Assert.IsNotNull(range.GetSurroundingContentControl());

            range.MoveEnd(Wd.WdUnits.wdCharacter, 1);
            Assert.IsNotNull(range.GetSurroundingContentControl());
        }

        [Test]
        public void Return_the_last_control_if_range_includes_many_content_controls()
        {
            Wd.Range range = this.doc.Range();
            Assert.IsNotNull(range.GetSurroundingContentControl());
        }

        [Test]
        public void Get_range_of_controls_without_values_and_insert_paragraph_at_end()
        {
            Wd.Range range = this.doc.Range();
            range.Start = this.doc.Range().ContentControls.First().Range.Start;
            range.End = this.doc.Range().ContentControls[4].Range.End;
            range.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);

            if (range.InContentControlOrContainsControls())
                range.MoveOutOfContentControl();
            range.InsertParagraphAfter();
        }

        [Test]
        public void Get_range_of_controls_in_a_table_without_values_and_insert_paragraph_at_end()
        {
            Wd.Range range = this.doc.Tables[1].Cell(1,2).Range;
            range.Start = range.ContentControls.First().Range.Start;
            range.End = range.ContentControls.Last().Range.End;
            range.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);

            if (range.InContentControlOrContainsControls())
                range.MoveOutOfContentControl();

            range.InsertParagraphAfter();
        }

        [Test]
        public void Get_range_of_controls_with_values_insert_paragraph_at_end()
        {
            Wd.Range range = this.doc.Range();
            range.Start = this.doc.ContentControls[5].Range.Start;
            range.End = this.doc.ContentControls[7].Range.End;
            range.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);

            if (range.InContentControlOrContainsControls())
                range.MoveOutOfContentControl();

            range.InsertParagraphAfter();
        }

    }
}
