using GorillaDocs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using O = Microsoft.Office.Core;
using XL = Microsoft.Office.Interop.Excel;

namespace GorillaDocs.Excel
{
    [ComVisible(true)]
    public abstract class BaseRibbon : O.IRibbonExtensibility
    {
        protected O.IRibbonUI ribbon;
        protected abstract XL.Application Application { get; }
        protected abstract string Resource { get; }
        protected abstract Assembly Assembly { get; }

        public void Ribbon_Load(O.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            Application.WorkbookActivate += app_WorkbookActivate;
        }

        void app_WorkbookActivate(XL.Workbook Wb) { ribbon.Invalidate(); }

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
    }
}
