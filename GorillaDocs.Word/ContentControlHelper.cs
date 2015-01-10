using GorillaDocs.libs.PostSharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    [Log]
    public static class ContentControlHelper
    {
        static List<Wd.ContentControl> GetContentControls(Wd.ContentControls ContentControls)
        {
            var contentControls = new List<Wd.ContentControl>();
            foreach (Wd.ContentControl control in ContentControls)
                contentControls.Add(control);
            return contentControls;
        }

        public static Wd.ContentControl[] FindAll(this Wd.ContentControls controls, string Tag)
        {
            return GetContentControls(controls).FindAll(x => x.Tag == Tag).ToArray();
        }
        public static Wd.ContentControl[] FindAllX(this Wd.ContentControls controls, string TagPattern)
        {
            return GetContentControls(controls).FindAll(x => Regex.IsMatch(x.Tag, TagPattern)).ToArray();
        }
        public static Wd.ContentControl[] FindAllByMappingX(this Wd.ContentControls controls, string MappingPattern)
        {
            return GetContentControls(controls).FindAll(x => x.XMLMapping != null && Regex.IsMatch(x.XMLMapping.XPath, MappingPattern)).ToArray();
        }
        public static Wd.ContentControl[] FindAllByNamespace(this Wd.ContentControls controls, string Namespace)
        {
            return GetContentControls(controls).FindAll(x => x.XMLMapping != null && x.XMLMapping.CustomXMLPart != null && x.XMLMapping.CustomXMLPart.NamespaceURI == Namespace).ToArray();
        }

        public static Wd.ContentControl Find(this Wd.ContentControls controls, string Tag)
        {
            return GetContentControls(controls).Find(x => x.Tag == Tag);
        }
        public static Wd.ContentControl FindX(this Wd.ContentControls controls, string TagPattern)
        {
            return GetContentControls(controls).Find(x => Regex.IsMatch(x.Tag, TagPattern));
        }

        public static Wd.ContentControl Add_Safely(this Wd.ContentControls controls, Wd.WdContentControlType type = Wd.WdContentControlType.wdContentControlRichText)
        {
            Handle_stupid_Word_bug_where_error_occurs_if_selection_is_in_control(controls);
            return controls.Add(type);
        }

        static void Handle_stupid_Word_bug_where_error_occurs_if_selection_is_in_control(Wd.ContentControls controls)
        {
            var selection = controls.Application.Selection;
            if (selection.Range.InContentControlOrContainsControls())
                selection.Range.MoveOutOfContentControl().Select();
        }

        public static bool InContentControl(this Wd.Range range)
        {
            Wd.Document doc = range.Parent;
            Wd.ContentControl control = doc.GetControlInRange(range);
            return control != null;
        }

        public static bool InContentControlOrContainsControls(this Wd.Range range)
        {
            return range.ContentControls.Count > 0 || range.InContentControl();
        }

        public static Wd.Range MoveOutOfContentControl(this Wd.Range range, Wd.WdCollapseDirection collapse = Wd.WdCollapseDirection.wdCollapseEnd)
        {
            Wd.ContentControl control = range.GetSurroundingContentControl();
            if (control != null)
                if (collapse == Wd.WdCollapseDirection.wdCollapseStart)
                {
                    if (range.Start >= control.Range.Start)
                        range.Start = control.Range.Start - 1;
                    if (range.End >= control.Range.Start)
                        range.End = control.Range.Start - 1;
                }
                else
                {
                    range.Start = control.Range.End;
                    if (range.Start <= control.Range.End)
                        range.Start = control.Range.End + 1;
                    if (range.End <= control.Range.End)
                        range.End = control.Range.End + 1;
                }
            if (range.ContainsTableCell())
                range.MoveEnd(Wd.WdUnits.wdCharacter, -1);

            //if (range.InContentControlOrContainsControls()) // Then it's because the Content Control is the first thing in the document
            //{
            //    range.Application.Selection.HomeKey(Wd.WdUnits.wdStory);
            //    range = range.Application.Selection.Range;
            //}
            return range;
        }

        public static Wd.ContentControl First(this Wd.ContentControls controls)
        {
            return controls[1];
        }
        public static Wd.ContentControl Last(this Wd.ContentControls controls)
        {
            return controls[controls.Count];
        }

        public static Wd.ContentControl GetSurroundingContentControl(this Wd.Range range)
        {
            Wd.Document doc = range.Parent;
            var controlCount = range.ContentControls.Count;
            if (controlCount > 0)
                return range.ContentControls[controlCount];
            else
                return GetControlInRange(doc, range);
        }

        public static Wd.ContentControl GetControlInRange(this Wd.Document doc, Wd.Range range)
        {
            // Use Temporary rage so that range is not modified by this routine.
            Wd.Range temp = range.Duplicate;
            Wd.Range selection = doc.Application.Selection.Range;
            try
            {
                ModifyRangeIfCollapsedBecauseTestBelowOnlyWorksIfControlIsEmpty(temp);
                temp.Select(); // Word Bug: The condition below does not always work if the range has not first been selected.
                foreach (Wd.ContentControl control in doc.ContentControls)
                    if (doc.Application.Selection.Range.InRange(control.Range))
                        return control;
                return null;
            }
            finally
            {
                selection.Select();
            }
        }

        static void ModifyRangeIfCollapsedBecauseTestBelowOnlyWorksIfControlIsEmpty(Wd.Range range)
        {
            if (range.Start == range.End)
                range.MoveStart(Wd.WdUnits.wdCharacter, -1);
        }

        public static void DeleteEmpty(this Wd.ContentControls controls)
        {
            foreach (Wd.ContentControl control in controls)
                if (string.IsNullOrEmpty(control.Range.Text) || control.ShowingPlaceholderText)
                    control.DeleteLine();
        }

        public static void FormatDates(this Wd.ContentControls controls, CultureInfo culture, string LongDateFormat, string expectedValue = null)
        {
            foreach (Wd.ContentControl control in controls)
                if (control.Type == Wd.WdContentControlType.wdContentControlDate)
                    FormatDateControl(LongDateFormat, culture, control, expectedValue);
        }
        public static void FormatDates(this Wd.ContentControls controls, CultureInfo culture, string LongDateFormat, string Tag, string expectedValue = null)
        {
            foreach (Wd.ContentControl control in controls)
                if (control.Type == Wd.WdContentControlType.wdContentControlDate)
                    if (control.Tag == Tag)
                        FormatDateControl(LongDateFormat, culture, control, expectedValue);
        }

        static BackgroundWorker worker;
        static void FormatDateControl(string LongDateFormat, CultureInfo culture, Wd.ContentControl control, string expectedValue = null)
        {
            try
            {
                control.DateDisplayLocale = (Wd.WdLanguageID)culture.LCID;
                control.DateDisplayFormat = LongDateFormat;
            }
            catch
            {
                // Not sure why, but sometimes error occurs when setting value from Content Control event. The worker should be fine, because it is not in the event.
            }

            if (!string.IsNullOrEmpty(expectedValue) && !string.Equals(control.Range.Text, expectedValue))
            {
                worker = new BackgroundWorker();
                worker.DoWork += WaitUntilThreadFreesDateControlThenSetFormat;
                worker.RunWorkerAsync(new List<object> { control, LongDateFormat, expectedValue });
            }
        }
        static void WaitUntilThreadFreesDateControlThenSetFormat(object sender, DoWorkEventArgs e)
        {
            try
            {
                int i = 0;
                List<object> args = (List<object>)e.Argument;
                Wd.ContentControl control = (Wd.ContentControl)args[0];
                string dateFormat = (string)args[1];
                string expectedValue = (string)args[2];
                while (i < 5 && !string.Equals(control.Range.Text, expectedValue))
                {
                    Thread.Sleep(250);
                    string tmp = control.XMLMapping.CustomXMLNode.Text;
                    if (tmp == "[Update]")
                        continue; // Other thread has set value to [Update] so continue.
                    control.XMLMapping.CustomXMLNode.Text = "[Update]";
                    control.XMLMapping.CustomXMLNode.Text = tmp;
                    control.DateDisplayFormat = dateFormat;
                    i++;
                }
            }
            catch (Exception ex)
            {
                // Log and ignore - the control may have been deleted by the time this code runs
                Message.LogWarning(ex);
            }
        }

        public static List<string> Unlock(this Wd.ContentControls controls)
        {
            var unlockedControls = new List<string>();
            foreach (Wd.ContentControl control in controls)
                if (control.LockContents)
                {
                    unlockedControls.Add(control.ID);
                    control.LockContents = false;
                }
            return unlockedControls;
        }

        public static void Add(this Wd.ContentControlListEntries ListEntries, IList<string> items)
        {
            ListEntries.Clear();
            foreach (string item in items)
                if (!string.IsNullOrEmpty(item))
                    ListEntries.Add(item, item);
        }

        public static bool IsNamed(this Wd.ContentControl control, string name)
        {
            return Regex.IsMatch(control.Tag, name);
        }

        public static void Delete(this Wd.ContentControl control, bool DeleteContents, string BookmarkName)
        {
            if (control != null)
            {
                Wd.Range range = control.Range;
                control.Delete(DeleteContents);
                range.Bookmarks.Add(BookmarkName, range);
            }
        }

        public static void DeleteControlAndSpace(this Wd.ContentControl control)
        {
            if (control != null)
            {
                Wd.Range range = control.Range;
                control.Delete(true);
                range.Delete();
            }
        }

        public static void DeleteRow(this Wd.ContentControl control)
        {
            control.Range.Rows[1].Delete();
        }

        public static void DeleteRows(this Wd.ContentControl control, int count = 1)
        {
            Wd.Range range = control.Range;
            for (int i = 0; i < count; i++)
                range.Rows[1].Delete();
        }

        public static void DeleteParagraph(this Wd.ContentControl control)
        {
            if (control != null)
                control.Range.Paragraphs[1].Range.Delete();
        }

        public static void DeleteParagraphIfEmpty(this Wd.ContentControl control)
        {
            Wd.Range range = control.Range;
            control.Delete(true);
            if (range.Paragraphs[1].IsEmpty())
                range.Paragraphs[1].Range.Delete();
        }

        public static void DeleteParagraph(this Wd.ContentControl control, string BookmarkName)
        {
            if (control != null)
            {
                Wd.Range range = control.Range;
                control.DeleteParagraph();
                range.Bookmarks.Add(BookmarkName, range);
            }
        }

        public static void DeleteLine(this Wd.ContentControl control)
        {
            var range = control.Range.Paragraphs[1].Range;
            if (range.Text.Contains("\a"))
            {
                range = control.Range;
                control.Delete(true);
                range.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);
                range.MoveStart(Wd.WdUnits.wdCharacter, -1);
                if (range.Text != null && !range.Text.Contains("\a"))
                    range.Delete();
            }
            else
                range.Delete();
        }

        public static void DeleteLine(this Wd.ContentControl control, string BookmarkName)
        {
            if (control != null)
            {
                Wd.Range range = control.Range;
                control.DeleteLine();
                range.Bookmarks.Add(BookmarkName, range);
            }
        }

        public static void DeleteAndTrim(this Wd.ContentControl control)
        {
            var range = control.Range.Paragraphs[1].Range;
            control.Delete(true);
            if (range.Characters.Count == 1)
                range.Delete();
        }

        public static string GetParentXPath(this Wd.ContentControl control)
        {
            string path = control.XMLMapping.XPath;
            path = path.Substring(0, path.LastIndexOf('/'));
            return path;
        }

        public static Wd.ContentControlListEntry SelectedItem(this Wd.ContentControlListEntries items)
        {
            foreach (Wd.ContentControlListEntry item in items)
            {
                var control = (Wd.ContentControl)items.Parent;
                if (item.Text.ToLower() == control.Range.Text.ToLower())
                    return item;
            }
            return null;
        }

        public static void DeleteParagraphAndRowIfEmpty(this Wd.ContentControl control)
        {
            Wd.Range range = control.Range;
            range.Paragraphs[1].Range.Delete();
            if (string.IsNullOrEmpty(range.Rows[1].Range.Text.Remove("\r\a")))
                range.Rows[1].Delete();
        }

        public static void UpdateLanguageId(this Wd.ContentControls controls, Wd.WdLanguageID LanguageID)
        {
            foreach (Wd.ContentControl control in controls)
            {
                control.Range.LanguageID = Wd.WdLanguageID.wdEnglishUS; // Set the default for Asian languages when typing in English
                control.Range.LanguageID = LanguageID;
            }
        }

        public static void SetIndex(this Wd.ContentControl control, int Index)
        {
            if (control != null)
                control.DropdownListEntries[Index].Select();
        }

        public static string GetValue(this Wd.ContentControl Control)
        {
            if (Control.Type == Wd.WdContentControlType.wdContentControlComboBox)
                foreach (Wd.ContentControlListEntry item in Control.DropdownListEntries)
                    if (item.Value.Equals(Control.Range.Text, StringComparison.OrdinalIgnoreCase))
                        return item.Value;
            return Control.Range.Text;
        }

        public static void DeleteUnMapped(this Wd.ContentControls controls)
        {
            var list = new List<Wd.ContentControl>();
            foreach (Wd.ContentControl control in controls)
                if (control.XMLMapping != null && !control.XMLMapping.IsMapped)
                    list.Add(control);
            foreach (Wd.ContentControl control in list)
                if (control.Range.Paragraphs[1].Range.ContentControls.Count > 1)
                    control.Delete(true);
                else
                    control.DeleteLine();
        }

        public static void DeleteEmptyMappedControls(this Wd.ContentControls controls)
        {
            // Have to delete in reverse order because of Word bug.
            // foreach worked fine in body of document
            // foreach does not work when controls are in header of document
            //foreach (Wd.ContentControl control in controls)
            for (int i = controls.Count; i > 0; i--)
            {
                var control = controls[i];
                if (control.XMLMapping.IsMapped && control.ShowingPlaceholderText == true)
                    control.Delete();
            }
        }

        public static bool IsMappedComboWithValueSelected(this Wd.ContentControl control)
        {
            return control.Type == Wd.WdContentControlType.wdContentControlComboBox && control.XMLMapping.IsMapped && control.ShowingPlaceholderText == false;
        }

        public static Wd.ContentControl SelectOrDefault(this Wd.ContentControls controls, int i)
        {
            try { return controls[i]; }
            catch { return null; }
        }

        public static bool ContainsSelection(this Wd.ContentControl control)
        {
            return control.Application.Selection.InRange(control.Range);
        }

        public static void ConvertMappedWithValueToText(this Wd.ContentControls controls)
        {
            foreach (Wd.ContentControl control in controls)
                if (control.XMLMapping.IsMapped && control.ShowingPlaceholderText == false)
                    control.Delete();
        }

        public static Wd.ContentControl MoveToNextContentControl(this Wd.Selection selection)
        {
            foreach (Wd.ContentControl control in selection.Document.ContentControls)
                if (control.Range.Start > selection.Range.End)
                {
                    control.Range.Select();
                    return control;
                }
            throw new InvalidOperationException("There are no more content controls.");
        }
        public static Wd.ContentControl MoveToPreviousContentControl(this Wd.Selection selection)
        {
            for (int i = selection.Document.ContentControls.Count; i > 0; i--)
            {
                Wd.ContentControl control = selection.Document.ContentControls[i];
                if (control.Range.End < selection.Range.Start)
                {
                    control.Range.Select();
                    return control;
                }
            }
            throw new InvalidOperationException("There are no more content controls.");
        }

        public static void ClearControlAndLock(this Wd.ContentControl control)
        {
            control.LockContents = false;
            control.SetPlaceholderText(null, null, "");
            control.Range.Text = "";
            control.LockContents = true;
        }

    }
}
