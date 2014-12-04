using GorillaDocs;
using System.IO;
using System.Reflection;

namespace GorillaDocs.IntegrationTests.Helpers
{
    public static class IOHelpers
    {
        public static FileInfo ContentControlsData { get { return Word_SampleData.GetFiles("ContentControls.docx")[0]; } }
        public static FileInfo FirmAddressData { get { return Word_SampleData.GetFiles("Firm Address.docx")[0]; } }
        public static FileInfo FirmAddressTestData { get { return Word_SampleData.GetFiles("Firm Address Test.docx")[0]; } }
        public static FileInfo EmptySectionsData { get { return Word_SampleData.GetFiles("Empty Sections.docx")[0]; } }
        public static FileInfo AppendixSection { get { return Word_Sections.GetFiles("Appendix.docx")[0]; } }
        public static FileInfo BackCoverSection { get { return Word_Sections.GetFiles("Back Cover.docx")[0]; } }
        public static FileInfo BlueDividerSection { get { return Word_Sections.GetFiles("Blue Divider.docx")[0]; } }
        public static FileInfo ExecutiveSummarySection { get { return Word_Sections.GetFiles("Executive Summary.docx")[0]; } }
        public static FileInfo FrontCoverSection { get { return Word_Sections.GetFiles("Front Cover.docx")[0]; } }
        public static FileInfo LandscapeSection { get { return Word_Sections.GetFiles("Landscape.docx")[0]; } }
        public static FileInfo PortraitSection { get { return Word_Sections.GetFiles("Portrait.docx")[0]; } }
        public static FileInfo TableOfContentsSection { get { return Word_Sections.GetFiles("Table of Contents.docx")[0]; } }
        public static FileInfo WhiteDividerSection { get { return Word_Sections.GetFiles("White Divider.docx")[0]; } }
        public static DirectoryInfo Word_SampleData
        {
            get
            {
                var folder = new DirectoryInfo(Assembly.GetExecutingAssembly().Path());
                folder = folder.Parent.Parent;
                folder = folder.GetDirectories("Word")[0];
                folder = folder.GetDirectories("SampleData")[0];
                return folder;
            }
        }
        public static DirectoryInfo Word_Sections { get { return Word_SampleData.GetDirectories("Sections")[0]; } }
    }
}
