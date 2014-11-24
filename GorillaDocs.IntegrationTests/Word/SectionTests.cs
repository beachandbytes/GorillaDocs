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
    public class SectionTests
    {
        // How to check that the tests are ok? other than visual inspection?
        Wd.Application wordApp;
        string TestSectionFile;

        [SetUp]
        public void setup()
        {
            wordApp = WordApplicationHelper.GetApplication();
            TestSectionFile = IOHelpers.BlueDividerSection.FullName;
        }

        //[Test]
        //public void Insert_Front_Cover_Section_at_the_start_of_the_document()
        //{
        //    Wd.Document doc = wordApp.Documents.Add(IOHelpers.EmptySectionsData.FullName);
        //    doc.Start().InsertSectionFromFile(IOHelpers.FrontCoverSection.FullName);
        //    doc.Sections[1].Delete();
        //    doc.Saved = true; // So that we don't get prompted to save when closing doc
        //}

        //[Test]
        //public void Insert_Test_Section_at_the_start_of_the_document()
        //{
        //    Wd.Document doc = wordApp.Documents.Add(IOHelpers.EmptySectionsData.FullName);
        //    doc.Start().InsertSectionFromFile(TestSectionFile);
        //    doc.Saved = true; // So that we don't get prompted to save when closing doc
        //}

        //[Test]
        //public void Insert_Test_Section_at_the_end_of_the_document()
        //{
        //    Wd.Document doc = wordApp.Documents.Add(IOHelpers.EmptySectionsData.FullName);
        //    doc.End().InsertSectionFromFile(TestSectionFile);
        //    doc.Saved = true; // So that we don't get prompted to save when closing doc
        //}

        //[Test]
        //public void Insert_Test_Section_at_the_start_of_a_section()
        //{
        //    Wd.Document doc = wordApp.Documents.Add(IOHelpers.EmptySectionsData.FullName);
        //    doc.Sections[2].Start().InsertSectionFromFile(TestSectionFile);
        //    doc.Saved = true; // So that we don't get prompted to save when closing doc
        //}

        //[Test]
        //public void Insert_Test_Section_in_the_middle_of_a_section()
        //{
        //    Wd.Document doc = wordApp.Documents.Add(IOHelpers.EmptySectionsData.FullName);
        //    doc.Paragraphs[2].Start().InsertSectionFromFile(TestSectionFile);
        //    doc.Saved = true; // So that we don't get prompted to save when closing doc
        //}

        //[Test]
        //public void Insert_Test_Section_at_the_end_of_a_section()
        //{
        //    Wd.Document doc = wordApp.Documents.Add(IOHelpers.EmptySectionsData.FullName);
        //    doc.Sections[1].End().InsertSectionFromFile(TestSectionFile);
        //    doc.Saved = true; // So that we don't get prompted to save when closing doc
        //}
    }
}
