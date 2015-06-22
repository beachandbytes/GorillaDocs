using GorillaDocs.libs.PostSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using O = Microsoft.Office.Core;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    [ComVisible(true)]
    public abstract class BaseRibbon : O.IRibbonExtensibility
    {
        protected O.IRibbonUI ribbon;
        protected abstract Wd.Application Application { get; }
        protected abstract string Resource { get; }
        protected abstract Assembly Assembly { get; }

        public virtual void Ribbon_Load(O.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            Application.DocumentChange += app_DocumentChange;
        }

        [System.Diagnostics.DebuggerStepThrough]
        void app_DocumentChange() { ribbon.Invalidate(); }

        [System.Diagnostics.DebuggerStepThrough]
        [QuietRibbonExceptionHandler]
        public bool IsDocOpen(O.IRibbonControl control) { return Application.ActiveDocumentExists(); }

        [System.Diagnostics.DebuggerStepThrough]
        [QuietRibbonExceptionHandler]
        public bool IsDocxCompatible(O.IRibbonControl control) { return Application.ActiveDocumentExists() ? (Application.ActiveDocument.CompatibilityMode >= 14) : false; }

        public string GetCustomUI(string RibbonID)
        {
            string ribbonXml = GetResourceText(Resource);
            ribbonXml = ribbonXml.Replace("%Version%", Assembly.FileVersion());
            return ribbonXml;
        }

        string GetResourceText(string resourceName)
        {
            string[] resourceNames = Assembly.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                    using (StreamReader resourceReader = new StreamReader(Assembly.GetManifestResourceStream(resourceNames[i])))
                        if (resourceReader != null)
                            return resourceReader.ReadToEnd();
            return null;
        }

        public virtual stdole.IPictureDisp GetImage(string imageName) { return ImageHelper.GetImage(imageName); }
    }
}
