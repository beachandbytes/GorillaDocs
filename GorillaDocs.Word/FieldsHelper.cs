using GorillaDocs.libs.PostSharp;
using System;
using O = Microsoft.Office.Core;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    [Log]
    public static class FieldsHelper
    {
        delegate void FieldActionDelegate(Wd.Field field);
        delegate void RangeActionDelegate(Wd.Range range);
        delegate void FieldTypeActionDelegate(Wd.Range range, Wd.WdFieldType type, string name);

        public static void UpdateAllFields(this Wd.Document doc)
        {
            doc.ProcessAllRanges(UpdateAllFields);
        }
        public static void UnlinkAllFields(this Wd.Document doc)
        {
            doc.ProcessAllRanges(UnlinkAllFields);
        }
        public static void UpdateFields(this Wd.Document doc, string name, Wd.WdFieldType type)
        {
            doc.ProcessAllRanges(type, name, UpdateFields);
        }
        public static void UnlinkFields(this Wd.Document doc, Wd.WdFieldType type, string name = null)
        {
            doc.ProcessAllRanges(type, name, UnlinkFields);
        }

        public static void UpdateAllFields(this Wd.Range range)
        {
            range.Fields.Update();
        }
        public static void UnlinkAllFields(this Wd.Range range)
        {
            range.Fields.Unlink();
        }
        public static void UpdateFields(this Wd.Range range, Wd.WdFieldType type, string name)
        {
            ProcessFields(range, type, name, UpdateField);
        }
        public static void UnlinkFields(this Wd.Range range, Wd.WdFieldType type, string name = null)
        {
            ProcessFields(range, type, name, UnlinkField);
        }
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

        public static bool ContainsField(this Wd.Range range, Wd.WdFieldType type, string name)
        {
            Wd.Range result = range.Find(type, name);
            return result.Find.Found;
        }
        public static void DeleteFields(this Wd.Range range, Wd.WdFieldType type, string name, bool DeleteEmptyParagraphs = true)
        {
            Wd.Range search = range.Duplicate;
            Wd.Range result = search.Find(type, name);
            while (result.Find.Found)
            {
                result.Delete();
                if (result.Paragraphs[1].Range.Characters.Count == 1 && DeleteEmptyParagraphs)
                    result.Paragraphs[1].Range.Delete();

                search.Start = result.End;
                result = search.Find(type, name);
            }
        }
        public static Wd.Field GetField(this Wd.Range range, Wd.WdFieldType type, string name)
        {
            Wd.Range search = range.Duplicate;
            Wd.Range result = search.Find(type, name);
            if (result.Find.Found)
                return result.Fields[1];
            else
                return null;
        }

        public static string AsGoToString(this Wd.WdFieldType type)
        {
            switch (type)
            {
                case Wd.WdFieldType.wdFieldDocProperty:
                    return "DOCPROPERTY";
                case Wd.WdFieldType.wdFieldDocVariable:
                    return "DOCVARIABLE";
                default:
                    throw new NotImplementedException();
            }
        }

        static void ProcessAllRanges(this Wd.Document doc, Wd.WdFieldType type, string name, FieldTypeActionDelegate action)
        {
            //TODO: Change to use StoryRanges

            action(doc.Range(), type, name);
            foreach (Wd.Section section in doc.Sections)
            {
                foreach (Wd.HeaderFooter header in section.Headers)
                    if (header.Exists)
                        action(header.Range, type, name);
                foreach (Wd.HeaderFooter footer in section.Footers)
                    if (footer.Exists)
                        action(footer.Range, type, name);
            }
            //NOTE: Can not use For Each loop on Bakers machines..
            //Works fine at MacroView. For some reason, infinite loop results at Bakers..
            //Change to standard For loop fixes Word bug.
            for (int i = 1; i <= doc.Shapes.Count; i++)
            {
                Wd.Shape shape = doc.Shapes[i];
                if (shape.Type == O.MsoShapeType.msoTextBox && shape.TextFrame.HasText == -1)
                    action(shape.TextFrame.TextRange, type, name);
            }
            //Update the textboxes in the headers - all the shapes from every section header and footer are available in the first header.
            for (int i = 1; i <= doc.Sections[1].Headers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count; i++)
            {
                Wd.Shape shape = doc.Sections[1].Headers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[i];
                if (shape.Type == O.MsoShapeType.msoTextBox && shape.TextFrame.HasText == -1)
                    action(shape.TextFrame.TextRange, type, name);
            }
            for (int i = 1; i <= doc.Sections[1].Footers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count; i++)
            {
                Wd.Shape shape = doc.Sections[1].Footers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[i];
                if (shape.Type == O.MsoShapeType.msoTextBox && shape.TextFrame.HasText == -1)
                    action(shape.TextFrame.TextRange, type, name);
            }
        }
        static void ProcessAllRanges(this Wd.Document doc, RangeActionDelegate action)
        {
            //TODO: Change to use StoryRanges

            action(doc.Range());
            foreach (Wd.Section section in doc.Sections)
            {
                foreach (Wd.HeaderFooter header in section.Headers)
                    if (header.Exists)
                        action(header.Range);
                foreach (Wd.HeaderFooter footer in section.Footers)
                    if (footer.Exists)
                        action(footer.Range);
            }
            //NOTE: Can not use For Each loop on Bakers machines..
            //Works fine at MacroView. For some reason, infinite loop results at Bakers..
            //Change to standard For loop fixes Word bug.
            for (int i = 1; i <= doc.Shapes.Count; i++)
            {
                Wd.Shape shape = doc.Shapes[i];
                if (shape.Type == O.MsoShapeType.msoTextBox && shape.TextFrame.HasText == -1)
                    action(shape.TextFrame.TextRange);
            }
            //Update the textboxes in the headers - all the shapes from every section header and footer are available in the first header.
            for (int i = 1; i <= doc.Sections[1].Headers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count; i++)
            {
                Wd.Shape shape = doc.Sections[1].Headers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[i];
                if (shape.Type == O.MsoShapeType.msoTextBox && shape.TextFrame.HasText == -1)
                    action(shape.TextFrame.TextRange);
            }
            for (int i = 1; i <= doc.Sections[1].Footers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count; i++)
            {
                Wd.Shape shape = doc.Sections[1].Footers[Wd.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[i];
                if (shape.Type == O.MsoShapeType.msoTextBox && shape.TextFrame.HasText == -1)
                    action(shape.TextFrame.TextRange);
            }
        }
        static void ProcessFields(Wd.Range range, Wd.WdFieldType type, string name, FieldActionDelegate action)
        {
            Wd.Range search = range.Duplicate;
            Wd.Range result = search.Find(type, name);
            while (result.Find.Found)
            {
                action(result.Fields[1]);
                search.Start = result.End;
                result = search.Find(type, name);
            }
        }
        static void UpdateField(Wd.Field field) { field.Update(); }
        static void UnlinkField(Wd.Field field) { field.Unlink(); }

    }
}
