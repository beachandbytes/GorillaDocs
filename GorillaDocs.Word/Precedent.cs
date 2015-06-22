using GorillaDocs.libs.PostSharp;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    [Log]
    public class Precedent<D>
    {
        public event EventHandler UpdateData;

        protected readonly Wd.Document Doc;
        readonly string NameSpace;
        readonly TaskScheduler taskScheduler;
        readonly Dispatcher dispatcher = null;
        bool DisableEvents = false;

        public Precedent(Wd.Document Doc, string NameSpace, bool MonitorContentControlEvents = true)
        {
            this.Doc = Doc;
            this.NameSpace = NameSpace;
            if (MonitorContentControlEvents)
                this.Doc.ContentControlBeforeContentUpdate += Doc_ContentControlBeforeContentUpdate;
            dispatcher = Dispatcher.CurrentDispatcher;
            SynchronizationContext.SetSynchronizationContext(new SynchronizationContext());
            taskScheduler = TaskScheduler.FromCurrentSynchronizationContext();
        }

        [System.Diagnostics.DebuggerStepThrough]
        // TODO: Figure out why deleted object errors occur in this event.
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
                dispatcher.Invoke(new Action(() =>
                {
                    try
                    {
                        if (Doc.IsTemplate()) return;

                        DisableEvents = true;
                        if (IsMonitoredControl(control))
                        {
                            if (UpdateData != null)
                                UpdateData(this, EventArgs.Empty);
                            // Changing this combo usually causes other mapped controls to update 
                            // so multiple Doc_ContentControlBeforeContentUpdate event may be running 
                            // If this combo causes a control to be deleted then errors may occur in other events
                            ProcessControls(Doc.ContentControls);
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

        public virtual bool IsMonitoredControl(Wd.ContentControl control) { return false; }

        public D Data { get; set; }
        public string VariableNameUsedInExpression { get; set; }

        public void ProcessControls() { ProcessControls(Doc.ContentControls); }

        public void ProcessControls_AfterCurrentEventCompletes(Wd.ContentControls controls)
        {
            SynchronizationContext.SetSynchronizationContext(new SynchronizationContext());

            Task T = Task.Factory.StartNew(() =>
            {
                System.Windows.Forms.Application.DoEvents(); // Let Word catch up and finish the Event.
            });

            T.ContinueWith((antecedent) =>
            {
                dispatcher.Invoke(new Action(() =>
                {
                    try
                    {
                        ProcessControls(controls);
                    }
                    catch (Exception ex)
                    {
                        Message.LogError(ex);
                    }
                }));
            }, taskScheduler);
        }

        public void ProcessControls(Wd.ContentControls controls)
        {
            for (int i = controls.Count; i > 0; i--)
            {
                var control = controls.SelectOrDefault(i);
                if (control != null)
                    ProcessControl(control);
            }
        }

        public virtual void ProcessControl(Wd.ContentControl control)
        {
            if (IsOptional(control))
                ProcessOptional(control);
            else if (IsDeleteRowIf(control))
                ProcessDeleteRowIf(control);
            else if (IsDeleteColumnIf(control))
                ProcessDeleteColumnIf(control);
            else if (IsClearCellIf(control))
                ProcessClearCellIf(control);
        }

        void ProcessOptional(Wd.ContentControl control)
        {
            var result = PrecedentExpression.Resolve(control.Tag, Data, VariableNameUsedInExpression);
            if (result == true)
            {
                try
                {
                    DeleteInstructions(control);
                    //control.Range.ContentControls.ConvertMappedWithValueToText();
                    control.Delete();
                }
                catch (Exception ex)
                {
                    Message.LogError(ex);
                }
            }
            else
                DeleteParagraphIfEmpty1(control);
        }

        void ProcessDeleteRowIf(Wd.ContentControl control)
        {
            var result = PrecedentExpression.Resolve(control.Tag, Data, VariableNameUsedInExpression);
            if (result)
                control.DeleteRow();
            else
                control.Delete();
        }

        void ProcessDeleteColumnIf(Wd.ContentControl control)
        {
            var result = PrecedentExpression.Resolve(control.Tag, Data, VariableNameUsedInExpression);
            if (result)
                control.DeleteColumnAndAutoFit();
            else
                control.Delete();
        }

        void ProcessClearCellIf(Wd.ContentControl control)
        {
            var result = PrecedentExpression.Resolve(control.Tag, Data, VariableNameUsedInExpression);
            if (result)
                control.ClearCell();
            else
                control.Delete();
        }

        public static void DeleteParagraphIfEmpty1(Wd.ContentControl control)
        {
            Wd.Range range = control.Range;
            control.Delete(true);
            if (range.Paragraphs[1].IsEmpty())
                range.Paragraphs[1].Range.Delete();
        }

        static bool IsOptional(Wd.ContentControl control)
        {
            return control.Type == Wd.WdContentControlType.wdContentControlRichText &&
                control.Title == "Optional";
        }

        static bool IsDeleteRowIf(Wd.ContentControl control)
        {
            return control.Type == Wd.WdContentControlType.wdContentControlRichText &&
                   control.Title == "DeleteRowIf";
        }

        static bool IsDeleteColumnIf(Wd.ContentControl control)
        {
            return control.Type == Wd.WdContentControlType.wdContentControlRichText &&
                   control.Title == "DeleteColumnIf";
        }

        static bool IsClearCellIf(Wd.ContentControl control)
        {
            return control.Type == Wd.WdContentControlType.wdContentControlRichText &&
                   control.Title == "ClearCellIf";
        }

        static void DeleteInstructions(Wd.ContentControl control)
        {
            try
            {
                foreach (Wd.ContentControl item in control.Range.ContentControls)
                    if (!string.IsNullOrEmpty(item.Title) && item.Title.ToLower() == "instruction")
                        item.DeleteParagraphIfEmpty();
            }
            catch (Exception ex)
            {
                Message.LogError(ex);
            }
        }
    }
}
