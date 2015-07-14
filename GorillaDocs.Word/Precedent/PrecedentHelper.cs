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
            else if (DeleteRowIf.Test(control))
                return new DeleteRowIf(control);
            else if (DeleteRowIf_RemoveControl.Test(control))
                return new DeleteRowIf_RemoveControl(control);
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

                PrecedentInstruction details = null;
                if (string.IsNullOrEmpty(control.Tag))
                    details = new PrecedentInstruction();
                else
                    details = Serializer.DeSerializeFromString<PrecedentInstruction>(control.Tag);

                var viewModel = new PrecedentInstructionViewModel(details, control, control.Range.Document.CustomXMLParts.Namespaces());
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

        public static void SetPrecedentInstruction(this Wd.ContentControl control, PrecedentInstruction instruction)
        {
            control.Tag = Serializer.SerializeToString<PrecedentInstruction>(instruction);
        }

        public static IList<string> Namespaces(this O.CustomXMLParts parts)
        {
            var namespaces = new List<string>();
            foreach (O.CustomXMLPart part in parts)
                namespaces.Add(part.NamespaceURI);
            return namespaces;
        }

        [System.Diagnostics.DebuggerStepThrough]
        public static PrecedentInstruction PrecedentInstruction(this Wd.ContentControl control)
        {
            try
            {
                if (!(control.Tag.StartsWith("<") && control.Tag.EndsWith(">")))
                    return null;
                else
                    return Serializer.DeSerializeFromString<PrecedentInstruction>(control.Tag);
            }
            catch
            {
                return null;
            }
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
