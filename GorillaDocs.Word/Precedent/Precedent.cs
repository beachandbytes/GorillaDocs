using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.Precedent
{
    //TODO: Bust these methods out into their own classes
    public class Precedent<D>
    {
        #region Constructor and Setup

        public event EventHandler UpdateData;

        protected readonly Wd.Document Doc;
        readonly TaskScheduler taskScheduler;
        readonly Dispatcher dispatcher = null;
        bool DisableEvents = false;

        public Precedent(Wd.Document Doc, bool MonitorContentControlEvents = true)
        {
            this.Doc = Doc;
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
                            ProcessControls(Doc.Range());
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

        #endregion

        #region Core

        public void ProcessControls_AfterCurrentEventCompletes(Wd.Range range)
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
                        ProcessControls(range);
                    }
                    catch (Exception ex)
                    {
                        Message.LogError(ex);
                    }
                }));
            }, taskScheduler);
        }

        public void ProcessControls(Wd.Range ProcessingRange = null)
        {
            if (ProcessingRange == null)
                ProcessingRange = Doc.Range();

            var controls = ProcessingRange.ContentControls;
            if (controls.Count > 0)
            {
                var control = controls.AsIList()[0];
                while (control != null)
                {
                    var range = control.Range;
                    ProcessControl(control);
                    control = range.MoveToNextControl(ProcessingRange);
                }
            }
        }

        public virtual void ProcessControl(Wd.ContentControl control)
        {
            var c = control.AsPrecedentControl();
            if (c != null)
            {
                c.Process();
                if (IsComboBox(control))
                    control.DropdownListEntries.Add(control.GetPrecedentInstruction().GetListItems(Doc));
            }
        }

        #endregion

        #region Helpers

        static bool IsComboBox(Wd.ContentControl control)
        {
            return control.Exists() && control.Type == Wd.WdContentControlType.wdContentControlComboBox && !string.IsNullOrEmpty(control.GetPrecedentInstruction().ListItems);
        }
        #endregion
    }
}
