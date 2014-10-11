using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.Xml.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public class Precedent
    {
        readonly Wd.Document Doc;
        readonly string NameSpace;
        readonly TaskScheduler taskScheduler;
        readonly Dispatcher dispatcher = null;
        bool DisableEvents = false;

        public Precedent(Wd.Document Doc, string NameSpace)
        {
            this.Doc = Doc;
            this.NameSpace = NameSpace;
            this.Doc.ContentControlBeforeContentUpdate += Doc_ContentControlBeforeContentUpdate;
            dispatcher = Dispatcher.CurrentDispatcher;
            SynchronizationContext.SetSynchronizationContext(new SynchronizationContext());
            taskScheduler = TaskScheduler.FromCurrentSynchronizationContext();
        }

        void Doc_ContentControlBeforeContentUpdate(Wd.ContentControl control, ref string Content)
        {
            if (DisableEvents) return;
            SynchronizationContext.SetSynchronizationContext(new SynchronizationContext());

            Task T = Task.Factory.StartNew(() =>
            {
                System.Windows.Forms.Application.DoEvents(); // Let Word catch up and finish the Event.
            });

            T.ContinueWith((antecedent) =>
            {
                if (Doc.IsTemplate()) return;

                dispatcher.Invoke(new Action(() =>
                {
                    try
                    {
                        DisableEvents = true;
                        if (control.IsMappedComboWithValueSelected())
                        {
                            // Changing this combo usually causes other mapped controls to update 
                            // so multiple Doc_ContentControlBeforeContentUpdate event may be running 
                            // If this combo causes a control to be deleted then errors may occur in other events
                            ProcessControls(Doc.ContentControls, Doc.CustomXmlPart(NameSpace).AsXDocument());
                            control.Delete();
                        }
                    }
                    catch (Exception ex)
                    {
                        Message.LogError(ex);
                    }
                    finally
                    {
                        DisableEvents = false;
                    }
                }));
            }, taskScheduler);
        }

        static void ProcessControls(Wd.ContentControls controls, XDocument data)
        {
            for (int i = controls.Count; i > 0; i--)
            {
                var control = controls.SelectOrDefault(i);
                if (control != null && IsOptional(control))
                {
                    var result = new OptionalCondition(control.Tag, data).Evaluate();
                    if (result == true)
                    {
                        try
                        {
                            DeleteInstructions(control);
                            control.Range.ContentControls.ConvertMappedWithValueToText();
                            control.Delete();
                        }
                        catch (Exception ex)
                        {
                            Message.LogError(ex);
                        }
                    }
                    else if (result == false)
                        control.DeleteParagraphIfEmpty();
                    // else null returned so we leave the control alone.
                }
            }
        }

        static bool IsOptional(Wd.ContentControl control)
        {
            bool newVariable = control.Type == Wd.WdContentControlType.wdContentControlRichText;
            bool newVariable1 = control.Title == "Optional";
            return newVariable && newVariable1;
        }

        static void DeleteInstructions(Wd.ContentControl control)
        {
            try
            {
                foreach (Wd.ContentControl item in control.Range.ContentControls)
                    if (item.Title.ToLower() == "instruction")
                        item.DeleteParagraphIfEmpty();
            }
            catch (Exception ex)
            {
                Message.LogError(ex);
            }
        }
    }
}
