using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.Precedent.Controls
{
    class RepeatingControl : PrecedentControl
    {
        public RepeatingControl(Wd.ContentControl control) : base(control) { }

        public static bool Test(Wd.ContentControl control)
        {
            return control.GetPrecedentInstruction() != null &&
                control.GetPrecedentInstruction().Command == "RepeatingControl";
        }

        public override void Process()
        {
            //TODO: Fix error when Repeating Control was originally set to the entire cell contents. Only works if CC control initially included paragraphs not end of cell character
            dynamic contacts = GetCollection(control);
            var title = control.Title;
            var controls = doc.Range().ContentControls(x => x.Title == title);

            if (contacts.Count == 0)
                DeleteControl(control);
            else
                while (contacts.Count != controls.Count)
                {
                    if (contacts.Count > controls.Count)
                    {
                        controls.Last().Copy();
                        var range = controls.Last().Range.CollapseEnd();
                        range.Move(Wd.WdUnits.wdCharacter, 1);
                        range.InsertParagraphBefore();
                        range.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);
                        range.Paste();
                        UpdateControls(range, controls.Count + 1, control.GetPrecedentInstruction().Expression);
                        range.Characters.Last.Delete();
                    }
                    else
                        DeleteControl(doc.Range().ContentControls(x => x.Title == title).Last());
                    controls = doc.Range().ContentControls(x => x.Title == title);
                }
        }

        static void DeleteControl(Wd.ContentControl control)
        {
            if (control.Range.Information[Wd.WdInformation.wdWithInTable])
            {
                if (control.Range == control.Range.Cells[1].Range)
                    control.Range.Rows[1].Delete();
                else
                    control.Range.ExpandParagraph().Delete();
            }
            else
                control.Range.ExpandParagraph().Delete();
        }

        dynamic GetCollection(Wd.ContentControl control)
        {
            var expression = control.GetPrecedentInstruction().Expression;
            var data = control.GetPrecedentInstruction().ExpressionData(control.Range.Document);

            var propertyInfo = data.GetType().GetProperty(expression);
            return propertyInfo.GetValue(data, null);
        }

        static void UpdateControls(Wd.Range range, int i, string CollectionName)
        {
            foreach (Wd.ContentControl control in range.ContentControls)
            {
                if (control.XMLMapping != null && !string.IsNullOrEmpty(control.XMLMapping.XPath))
                    control.XMLMapping.SetMapping(UpdateContactIndex(control.XMLMapping.XPath, i));

                var instruction = control.GetPrecedentInstruction();
                if (instruction != null)
                {
                    instruction.Expression = Regex.Replace(instruction.Expression, string.Format(@"{0}\[[\d]*\]", CollectionName), string.Format("{0}[{1}]", CollectionName, i - 1));
                    if (instruction.ListItems != null)
                        instruction.ListItems = Regex.Replace(instruction.ListItems, string.Format(@"{0}\[[\d]*\]", CollectionName), string.Format("{0}[{1}]", CollectionName, i - 1));
                    control.SetPrecedentInstruction(instruction);
                }
            }
        }

        static string UpdateContactIndex(string xPath, int index)
        {
            //TODO: How to unhardcode 'Contact' and 'Party' without adding it to PrecedentInstruction?
            xPath = Regex.Replace(xPath, @"Contact\[[\d]*\]", string.Format("Contact[{0}]", index));
            xPath = Regex.Replace(xPath, @"Party\[[\d]*\]", string.Format("Party[{0}]", index));
            return xPath;
        }

    }
}
