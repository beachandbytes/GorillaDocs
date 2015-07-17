using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace GorillaDocs.OpenXml
{
    public class WordDocument
    {
        readonly FileInfo file;
        public WordDocument(FileInfo file) { this.file = file; }

        public IList<Variable> Variables
        {
            get
            {
                var variables = new List<Variable>();
                using (Stream docStream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var doc = WordprocessingDocument.Open(docStream, false))
                    foreach (DocumentVariables docVars in doc.MainDocumentPart.DocumentSettingsPart.Settings.Descendants<DocumentVariables>().ToList())
                        foreach (DocumentVariable docVar in docVars)
                            variables.Add(new Variable() { Name = docVar.Name, Value = docVar.Val.Value });
                return variables;
            }
        }

        public XElement GetCustomXML(string ns)
        {
            using (Stream docStream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var doc = WordprocessingDocument.Open(docStream, false))
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
