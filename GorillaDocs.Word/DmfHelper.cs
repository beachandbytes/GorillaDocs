using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class DmfHelper
    {
        public static void SaveAsToDmf(this Wd.Document Doc)
        {
            var addin = Doc.Application.COMAddIns.Find("MacroView.DMF.Word");
            if (addin == null)
                throw new InvalidOperationException("MacroView DMF is not installed.");
            else
            {
                dynamic dmfService = addin.Object;
                dmfService.ShowSaveAsDialog();
            }
        }
    }
}
