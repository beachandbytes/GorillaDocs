using GorillaDocs.Views;
using GorillaDocs.Word.MacroView.Properties;
using MacroView.DMF.Office.Extensibility.Word;
using System.Collections;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.MacroView
{
    public static class DmfHelper
    {
        public static bool IsDmfInstalled(this Wd.Application app) { return app.COMAddIns.IsLoaded(Properties.Settings.Default.DMFWordAddin); }

        public static void DmfOpen(this Wd.Documents docs)
        {
            var service = (DmfDocumentAutomationService)docs.Application.COMAddIns.Find(Settings.Default.DMFWordAddin).Object;
            service.ShowOpenDialog();
        }

        public static void DmfSaveAs(this Wd.Document doc, Hashtable properties = null, bool StandardSaveOnCancel = true)
        {
            var view = new WaitingView();
            view.Show();
            try
            {
                var service = (DmfDocumentAutomationService)doc.Application.COMAddIns.Find(Settings.Default.DMFWordAddin).Object;
                if (!service.SaveAs(ref properties, null, true, false) && StandardSaveOnCancel)
                    doc.Application.Dialogs[Wd.WdWordDialog.wdDialogFileSaveAs].Show();
            }
            finally
            {
                view.Close();
            }
        }

        public static void DmfSaveAs(this Wd.Document Doc)
        {
            var view = new WaitingView();
            view.Show();
            try
            {
                dynamic dmfService = Doc.Application.COMAddIns.Find("MacroView.DMF.Word").Object;
                dmfService.ShowSaveAsDialog();
            }
            finally
            {
                view.Close();
            }
        }
    }
}
