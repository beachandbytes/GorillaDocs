using GorillaDocs;
using System.IO;
using System.Reflection;

namespace GorillaDocs.IntegrationTests.Helpers
{
    public static class IOHelpers
    {
        public static FileInfo ContentControlsData { get { return Word_SampleData.GetFiles("ContentControls.docx")[0]; } }
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
    }
}
