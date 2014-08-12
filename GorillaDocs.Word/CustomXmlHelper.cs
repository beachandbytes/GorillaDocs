using GorillaDocs.libs.PostSharp;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using O = Microsoft.Office.Core;

namespace GorillaDocs.Word
{
    [Log]
    public static class CustomXmlHelper
    {
        public static DateTime GetNodeDate(this O.CustomXMLPart part, string NodeName, CultureInfo culture)
        {
            try
            {
                DateTime result;
                var node = part.SelectSingleNode(string.Format("/ns0:root/ns0:{0}", NodeName));
                if (node == null || string.IsNullOrEmpty(node.Text))
                    return DateTime.Now;
                else if (DateTime.TryParse(node.Text, out result))
                    return result;
                else if (DateTime.TryParse(node.Text, culture, DateTimeStyles.None, out result))
                    return result;
                else
                    throw new InvalidOperationException(string.Format("Unable to convert '{0}' to DateTime value.", node.Text));
            }
            catch (Exception ex)
            {
                return DateTime.Now;
            }
        }
        public static string GetNodeText(this O.CustomXMLPart part, string NodeName)
        {
            var node = part.SelectSingleNode(string.Format("/ns0:root/ns0:{0}", NodeName));
            if (node == null)
                return string.Empty;
            else
                return node.Text;
        }
        public static string GetNodeText(this O.CustomXMLPart part, string ParentNodeName, string NodeName)
        {
            var node = part.SelectSingleNode(string.Format("/ns0:root/ns0:{0}", ParentNodeName));
            node = node.SelectSingleNode("ns0:" + NodeName);
            if (node == null)
                return string.Empty;
            else
                return node.Text;
        }
        public static string GetNodeText(this O.CustomXMLPart part, string ParentNodeName, string NodeName, int i)
        {
            var node = part.SelectNodes(string.Format("/ns0:root/ns0:{0}s/ns0:{0}", ParentNodeName))[i];
            node = node.SelectSingleNode("ns0:" + NodeName);
            if (node == null)
                return string.Empty;
            else
                return node.Text;
        }
        public static string GetNodeText(this O.CustomXMLNode node, string NodeName)
        {
            var childNode = node.SelectSingleNode(string.Format("ns0:{0}", NodeName));
            if (childNode == null)
                return string.Empty;
            else
                return childNode.Text;
        }
        public static void SetNodeDate(this O.CustomXMLPart part, string NodeName, DateTime? value)
        {
            if (value == null)
                return;
            var node = part.SelectSingleNode(string.Format("/ns0:root/ns0:{0}", NodeName));
            if (node == null)
                part.DocumentElement.AppendChildNode(NodeName, part.NamespaceURI, O.MsoCustomXMLNodeType.msoCustomXMLNodeElement, ((DateTime)value).ToString("s"));
            else
            {
                node.Text = ((DateTime)value).ToString("s");
                DateTime d;
                if (!DateTime.TryParse(node.Text, out d))
                    node.Text = ((DateTime)value).ToString("s"); // Sometimes need to set twice due to Word bug..
            }
        }
        public static void SetNodeText(this O.CustomXMLPart part, string NodeName, string value)
        {
            var node = part.SelectSingleNode(string.Format("/ns0:root/ns0:{0}", NodeName));
            if (node == null)
                part.DocumentElement.AppendChildNode(NodeName, part.NamespaceURI, O.MsoCustomXMLNodeType.msoCustomXMLNodeElement, value);
            else
                node.Text = value;
        }
        public static void SetNodeText(this O.CustomXMLPart part, string ParentNodeName, string NodeName, string value)
        {
            var parentNode = part.SelectSingleNode(string.Format("/ns0:root/ns0:{0}", ParentNodeName));
            var node = parentNode.SelectSingleNode("ns0:" + NodeName);
            if (node == null)
                parentNode.AppendChildNode(NodeName, part.NamespaceURI, O.MsoCustomXMLNodeType.msoCustomXMLNodeElement, value);
            else
                node.Text = value;
        }
        public static void SetNodeText(this O.CustomXMLPart part, string ParentNodeName, string NodeName, int i, string value)
        {
            var parentNode = part.SelectNodes(string.Format("/ns0:root/ns0:{0}s/ns0:{0}", ParentNodeName))[i];
            var node = parentNode.SelectSingleNode("ns0:" + NodeName);
            if (node == null)
                parentNode.AppendChildNode(NodeName, part.NamespaceURI, O.MsoCustomXMLNodeType.msoCustomXMLNodeElement, value);
            else
                node.Text = value;
        }
        public static void SetNodeText(this O.CustomXMLNode node, string NodeName, string value)
        {
            var childNode = node.SelectSingleNode(string.Format("ns0:{0}", NodeName));
            if (childNode == null)
                node.AppendChildNode(NodeName, node.NamespaceURI, O.MsoCustomXMLNodeType.msoCustomXMLNodeElement, value);
            else
                childNode.Text = value;
        }

        public static string GetAttribute(this O.CustomXMLNode node, string AttributeName)
        {
            foreach (O.CustomXMLNode attribute in node.Attributes)
                if (attribute.BaseName == AttributeName)
                    return attribute.NodeValue;
            return string.Empty;
        }
        public static void SetAttribute(this O.CustomXMLNode node, string AttributeName, string value)
        {
            foreach (O.CustomXMLNode attribute in node.Attributes)
                if (attribute.BaseName == AttributeName)
                    attribute.NodeValue = value;
        }

        public static void DeleteByNamespace(this O.CustomXMLParts parts, string Namespace)
        {
            foreach (O.CustomXMLPart customXMLPart in parts.SelectByNamespace(Namespace))
                customXMLPart.Delete();
        }
    }
}
