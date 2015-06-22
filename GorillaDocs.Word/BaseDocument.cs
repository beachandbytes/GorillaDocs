using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;
using O = Microsoft.Office.Core;

namespace GorillaDocs.Word
{
    public class BaseDocument
    {
        protected readonly Wd.Document doc = null;
        public virtual string NameSpace { get { return "http://schemas.macroview.com.au/document"; } }

        public BaseDocument(Wd.Document Doc) { doc = Doc; }

        public O.CustomXMLPart CustomXmlPart
        {
            get
            {
                var parts = doc.CustomXMLParts.SelectByNamespace(NameSpace);
                return parts.Count == 1 ? parts[1] : null;
            }
        }

        public Wd.Application Application { get { return this.doc.Application; } }
        public Wd.Selection Selection { get { return this.doc.Application.Selection; } }
        public void Activate() { this.doc.Activate(); }
        public void Close(Wd.WdSaveOptions SaveChanges) { this.doc.Close(SaveChanges); }
        public string Name { get { return this.doc.Name; } }
        public string Path { get { return this.doc.Path; } }
        public Wd.Range Range() { return this.doc.Range(); }
        public void Save() { this.doc.Save(); }
        public Wd.Bookmarks Bookmarks { get { return this.doc.Bookmarks; } }
        public Wd.Shapes Shapes { get { return this.doc.Shapes; } }
        public Wd.Styles Styles { get { return this.doc.Styles; } }
        public string FullName
        {
            [System.Diagnostics.DebuggerStepThrough]
            get
            {
                try
                {
                    return this.doc.FullName;
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
