using GorillaDocs.Word.Precedent;
using System.Globalization;
using System.Xml.Linq;
using O = Microsoft.Office.Core;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public class BaseDocument
    {
        protected readonly Wd.Document doc = null;
        public virtual string NameSpace { get { return "http://schemas.macroview.com.au/document"; } }

        public BaseDocument(Wd.Document Doc) { doc = Doc; }

        public virtual void NewDocument(ref bool Cancel) { }
        public virtual void EditDetails(ref bool Cancel) { }

        public virtual O.CustomXMLPart CustomXmlPart
        {
            get
            {
                var parts = doc.CustomXMLParts.SelectByNamespace(NameSpace);
                if (parts.Count == 1)
                    return parts[1];
                else
                {
                    var OpenXmlDoc = new OpenXml.WordDocument(doc.get_AttachedTemplate());
                    XElement xml = OpenXmlDoc.GetCustomXML(NameSpace);
                    if (xml != null)
                        return doc.CustomXMLParts.Add(xml.ToString());
                }
                return null;
            }
        }

        public void ProcessControls<T>(Wd.Range range = null)
        {
            var precedent = new Precedent<T>(doc);
            precedent.ProcessControls(range);
        }

        protected void RemoveEditDetailsBookmarks()
        {
            foreach (Wd.Bookmark bookmark in Bookmarks)
                if (bookmark.Name.StartsWith("EditDetails_"))
                    bookmark.Delete();
        }

        public virtual void ReplaceLanguageStrings() { }

        public void SetProofingLanguage(Wd.Range range, CultureInfo Culture)
        {
            var lcid = (Wd.WdLanguageID)Culture.LCID;
            range.LanguageID = lcid;
            range.ContentControls.UpdateLanguageId(lcid);
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
        public Wd.Sections Sections { get { return this.doc.Sections; } }
        public Wd.TablesOfContents TablesOfContents { get { return this.doc.TablesOfContents; } }
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
        public string CreatedVersion
        {
            get { return doc.GetDocVar("Version"); }
            set { doc.SetDocVar("Version", value); }
        }

        public virtual bool ProfileNewDocument() { return false; }
    }
}
