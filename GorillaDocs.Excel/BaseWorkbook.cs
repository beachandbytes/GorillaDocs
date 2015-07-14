using System;
using System.Collections.Generic;
using System.Linq;
using O = Microsoft.Office.Core;
using XL = Microsoft.Office.Interop.Excel;

namespace GorillaDocs.Excel
{
    public class BaseWorkbook
    {
        protected readonly XL.Workbook workbook = null;
        public virtual string NameSpace { get { return "http://schemas.macroview.com.au/workbook"; } }

        public BaseWorkbook(XL.Workbook Wb) { workbook = Wb; }

        public virtual void RegisterEvents() { }

        public O.CustomXMLPart CustomXmlPart
        {
            get
            {
                var parts = workbook.CustomXMLParts.SelectByNamespace(NameSpace);
                return parts.Count == 1 ? parts[1] : null;
            }
        }

        public XL.Application Application { get { return workbook.Application; } }
        public XL.Range Selection { get { return Application.Selection; } }
        public void Activate() { workbook.Activate(); }
        public string Name { get { return workbook.Name; } }
        public string Path { get { return workbook.Path; } }
        public void Save() { workbook.Save(); }
        public void Close(bool? SaveChanges) { workbook.Close(SaveChanges); }
        public string FullName
        {
            [System.Diagnostics.DebuggerStepThrough]
            get
            {
                try
                {
                    return workbook.FullName;
                }
                catch
                {
                    // This errors when document has been closed.
                    return string.Empty;
                }
            }
        }
    }
}
