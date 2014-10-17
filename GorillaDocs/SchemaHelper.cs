using System;
using System.IO;
using System.Xml.Serialization;

namespace GorillaDocs
{
    public static class SchemaHelper
    {
        public static void SaveSchema<T>(this T obj, string FullName)
        {
            var importer = new XmlReflectionImporter();
            var schemas = new XmlSchemas();
            var exporter = new XmlSchemaExporter(schemas);
            Type type = obj.GetType();
            XmlTypeMapping map = importer.ImportTypeMapping(type);
            exporter.ExportTypeMapping(map);

            TextWriter tw = new StreamWriter(FullName);
            schemas[0].Write(tw);
            tw.Close();
        }
    }
}
