using GorillaDocs.Word.Precedent.Controls;
using GorillaDocs.Word.Precedent.ViewModels;
using GorillaDocs.Word.Precedent.Views;
using System;
using System.Collections.Generic;
using O = Microsoft.Office.Core;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.Precedent
{
    public static class PrecedentHelper
    {
        public static PrecedentControl AsPrecedentControl(this Wd.ContentControl control)
        {
            if (Optional.Test(control))
                return new Optional(control);
            else if (DeleteControlIf.Test(control))
                return new DeleteControlIf(control);
            else if (DeleteLineIf.Test(control))
                return new DeleteLineIf(control);
            else if (DeleteLineIf_OrRemoveControl.Test(control))
                return new DeleteLineIf_OrRemoveControl(control);
            else if (DeleteRowIf.Test(control))
                return new DeleteRowIf(control);
            else if (DeleteRowIf_OrRemoveControl.Test(control))
                return new DeleteRowIf_OrRemoveControl(control);
            else if (DeleteColumnIf.Test(control))
                return new DeleteColumnIf(control);
            else if (ClearCellIf.Test(control))
                return new ClearCellIf(control);
            else if (RepeatingControl.Test(control))
                return new RepeatingControl(control);
            else if (ContactDelivery.Test(control))
                return new ContactDelivery(control);
            else
                return null;
        }

        public static void EditPrecedentInstruction(this Wd.ContentControl control)
        {
            try
            {
                control.Range.Select();
                var instruction = GetPrecedentInstruction(control);
                if (instruction == null)
                    instruction = new PrecedentInstruction();

                var viewModel = new PrecedentInstructionViewModel(instruction, control, control.Range.Document.CustomXMLParts.Namespaces());
                var view = new PrecedentInstructionView(viewModel);
                view.ShowDialog();
                if (view.DialogResult == true)
                {
                    control.SetPrecedentInstruction(view.viewModel.Details);
                    control.Range.Document.Saved = false;
                }
            }
            catch (Exception ex)
            {
                if (ex.InnerException != null && ex.InnerException is System.Xml.XmlException)
                    throw new InvalidOperationException("The control Tag property is not a valid Precedent Instruction.\nClear the Tag property and try again.");
            }
        }

        public static PrecedentInstruction GetPrecedentInstruction(this Wd.ContentControl control)
        {
            // Precedent Instruction saved as DocVar instead of Control.Tag because Word2010 has a 64 character limit..
            PrecedentInstruction details = null;
            var xml = control.Range.Document.GetDocVar("PrecedentInstruction_" + control.ID);

            if (!string.IsNullOrEmpty(xml))
                details = Serializer.DeSerializeFromString<PrecedentInstruction>(xml);
            else if (!string.IsNullOrEmpty(control.Tag) && control.Tag.StartsWith("PrecedentInstruction_"))
                details = Serializer.DeSerializeFromString<PrecedentInstruction>(control.Range.Document.GetDocVar(control.Tag));
            else if (!string.IsNullOrEmpty(control.Tag) && control.Tag.StartsWith("<") && control.Tag.EndsWith(">")) //TODO: Delete this case once working at Webb Henderson
            {
                details = Serializer.DeSerializeFromString<PrecedentInstruction>(control.Tag);
                control.Tag = ""; 
            }
            return details;
        }

        public static void SetPrecedentInstruction(this Wd.ContentControl control, PrecedentInstruction instruction)
        {
            // Precedent Instruction saved as DocVar instead of Control.Tag because Word2010 has a 64 character limit..
            control.Range.Document.SetDocVar("PrecedentInstruction_" + control.ID, Serializer.SerializeToString<PrecedentInstruction>(instruction));
            control.Tag = "PrecedentInstruction_" + control.ID; // So that we can copy the Precedent Instruction when a Repeating Contact is added.
        }

        public static IList<string> Namespaces(this O.CustomXMLParts parts)
        {
            var namespaces = new List<string>();
            foreach (O.CustomXMLPart part in parts)
                namespaces.Add(part.NamespaceURI);
            return namespaces;
        }

        public static Wd.ContentControl MoveToNextControl(this Wd.Range range, Wd.Range ProcessingRange)
        {
            var controls = ProcessingRange.ContentControls.AsIList();
            foreach (Wd.ContentControl control in controls)
                if (control.Exists() && control.Range.Start > range.Start)
                    return control;
            return null;
        }

        public static IList<Wd.ContentControl> AsIList(this Wd.ContentControls controls)
        {
            var list = new List<Wd.ContentControl>();
            foreach (Wd.ContentControl control in controls)
                list.Add(control);
            return list;
        }
    }
}
