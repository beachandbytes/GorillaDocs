using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace GorillaDocs
{
    public class Serializer
    {
        /// <summary>
        /// Serializes an object.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="serializableObject"></param>
        /// <param name="fileName"></param>
        public static void SerializeToFile<T>(T serializableObject, string fileName)
        {
            if (serializableObject == null) { return; }

            XmlSerializer serializer = new XmlSerializer(serializableObject.GetType());
            using (TextWriter textWriter = new StreamWriter(fileName))
                serializer.Serialize(textWriter, serializableObject);
        }
        public static string SerializeToString<T>(T serializableObject)
        {
            XmlSerializer serializer = new XmlSerializer(serializableObject.GetType());
            XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
            ns.Add("", "");
            using (TextWriter textWriter = new StringWriter())
            {
                serializer.Serialize(textWriter, serializableObject, ns);
                return textWriter.ToString();
            }
        }
        public static XDocument SerializeToXDocument<T>(T serializableObject)
        {
            var serializer = new DataContractSerializer(serializableObject.GetType());
            XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
            ns.Add("", "");
            XDocument doc = new XDocument();
            using (var writer = doc.CreateWriter())
            {
                serializer.WriteObject(writer, serializableObject);
                return doc;
            }
        }
        public static XmlDocument SerializeToXmlDocument<T>(T serializableObject)
        {
            XmlSerializer serializer = new XmlSerializer(serializableObject.GetType());
            XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
            ns.Add("", "");
            using (TextWriter textWriter = new StringWriter())
            {
                serializer.Serialize(textWriter, serializableObject, ns);
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(textWriter.ToString());
                return doc;
            }
        }

        /// <summary>
        /// Deserializes an xml file into an object list
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static T DeSerializeFromFile<T>(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) { return default(T); }

            XmlSerializer deserializer = new XmlSerializer(typeof(T));
            using (TextReader textReader = new StreamReader(fileName))
                return (T)deserializer.Deserialize(textReader);
        }
        public static T DeSerializeFromString<T>(string value)
        {
            XmlSerializer deserializer = new XmlSerializer(typeof(T));
            using (TextReader textReader = new StringReader(value))
                return (T)deserializer.Deserialize(textReader);
        }
        public static T DeSerializeFromXDocument<T>(XDocument value)
        {
            XmlSerializer deserializer = new XmlSerializer(typeof(T));
            using (TextReader textReader = new StringReader(value.ToString()))
                return (T)deserializer.Deserialize(textReader);
        }
        public static T DeSerializeFromXDocument<T>(XmlDocument value)
        {
            XmlSerializer deserializer = new XmlSerializer(typeof(T));
            using (TextReader textReader = new StringReader(value.ToString()))
                return (T)deserializer.Deserialize(textReader);
        }
    }
}
