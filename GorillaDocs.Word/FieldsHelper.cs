using GorillaDocs.libs.PostSharp;
using O = Microsoft.Office.Core;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    [Log]
    public static class FieldsHelper
    {
        public static void UpdateFields(this Wd.HeaderFooter headerfooter)
        {
            headerfooter.Range.Fields.Update();
            //Update the textboxes in the headers - all the shapes from every section header and footer are available in the first header.
            for (int i = 1; i <= headerfooter.Shapes.Count; i++)
            {
                Wd.Shape shape = headerfooter.Shapes[i];
                if (shape.Type == O.MsoShapeType.msoTextBox && shape.TextFrame.HasText == 1)
                    shape.TextFrame.TextRange.Fields.Update();
            }
        }

        public static void UpdateAllFields(this Wd.Document doc)
        {
            doc.Fields.Update();
            foreach (Wd.Section section in doc.Sections)
            {
                foreach (Wd.HeaderFooter header in section.Headers)
                    if (header.Exists)
                        header.Range.Fields.Update();
                foreach (Wd.HeaderFooter footer in section.Footers)
                    if (footer.Exists)
                        footer.Range.Fields.Update();
            }
            //NOTE: Can not use For Each loop ..
            //For some reason, infinite loop results on some machines..
            //Change to standard For loop fixes Word bug.
            for (int i = 1; i <= doc.Shapes.Count; i++)
            {
                Wd.Shape shape = doc.Shapes[i];
                if (shape.Type == O.MsoShapeType.msoTextBox && shape.TextFrame.HasText == -1)
                    shape.TextFrame.TextRange.Fields.Update();
            }
            //Update the textboxes in the headers - all the shapes from every section header and footer are available in the first header.
            for (int i = 1; i <= doc.Sections[1].Headers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count; i++)
            {
                Wd.Shape shape = doc.Sections[1].Headers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[i];
                if (shape.Type == O.MsoShapeType.msoTextBox && shape.TextFrame.HasText == -1)
                    shape.TextFrame.TextRange.Fields.Update();
            }
            for (int i = 1; i <= doc.Sections[1].Footers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count; i++)
            {
                Wd.Shape shape = doc.Sections[1].Footers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[i];
                if (shape.Type == O.MsoShapeType.msoTextBox && shape.TextFrame.HasText == -1)
                    shape.TextFrame.TextRange.Fields.Update();
            }
        }
    }
}
