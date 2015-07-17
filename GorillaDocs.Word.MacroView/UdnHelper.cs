using GorillaDocs.Word.MacroView.Properties;
using macroview.udn.Extensibility.Word;
using System.Collections.Generic;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.MacroView
{
    public static class UdnHelper
    {
        public static bool IsUdnInstalled(this Wd.Application app) { return app.COMAddIns.IsLoaded(Settings.Default.UDNWordAddin); }

        public static bool IsUdnDoc(this Wd.Document doc) { return doc.Path.ToLower().StartsWith("http") && !string.IsNullOrEmpty(doc.GetDocProp("mvRef")); }

        public static void ToggleReference(this BaseDocument doc)
        {
            if (doc.Sections.Footers().UdnReferences().Exists())
                doc.Sections.Footers().UdnReferences().Delete();
            else
                doc.Sections.Footers().UdnInsertReference();
        }

        public static IList<Wd.Field> UdnReferences(this IList<Wd.HeaderFooter> headerFooters)
        {
            var fields = new List<Wd.Field>();
            foreach (Wd.HeaderFooter headerFooter in headerFooters)
                foreach (Wd.Field field in headerFooter.Range.Fields)
                    if (field.Code.Text.Contains("mvRef"))
                        fields.Add(field);
            return fields;
        }

        public static void UdnInsertReference(this Wd.Selection selection)
        {
            var service = (UdnDocumentAutomationService)selection.Application.COMAddIns.Find(Settings.Default.UDNWordAddin).Object;
            service.InsertReference();
        }

        public static void UdnInsertReference(this Wd.Section section)
        {
            var doc = (Wd.Document)section.Parent;
            if (doc.IsUdnDoc() && doc.Sections.Footers().UdnReferences().Exists())
                section.Footers().UdnInsertReference();
        }

        public static void UdnInsertReference(this IList<Wd.HeaderFooter> headerFooters)
        {
            foreach (Wd.HeaderFooter headerFooter in headerFooters)
            {
                var range = headerFooter.Range.CollapseEnd();
                range.Fields.Add(range, Wd.WdFieldType.wdFieldDocProperty, "mvRef");
            }
        }

    }
}
