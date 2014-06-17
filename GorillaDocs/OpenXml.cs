using GorillaDocs.libs.PostSharp;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using OX = DocumentFormat.OpenXml;

namespace GorillaDocs
{
    [Log]
    public class OpenXml
    {
        public static XElement GetCustomXML(FileInfo file, string ns)
        {
            using (Stream docStream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var doc = OX.Packaging.WordprocessingDocument.Open(docStream, false))
                foreach (var part in doc.MainDocumentPart.CustomXmlParts)
                    using (var reader = new XmlTextReader(part.GetStream(FileMode.Open, FileAccess.Read)))
                    {
                        reader.MoveToContent();
                        if (reader.NamespaceURI == ns)
                            return XElement.Load(reader);
                    }
            return null;
        }
    }
}
