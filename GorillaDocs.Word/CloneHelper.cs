using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;
using O = Microsoft.Office.Core;
using GorillaDocs.libs.PostSharp;
using System.Reflection;

namespace GorillaDocs.Word
{
    [Log]
    public static class CloneHelper
    {
        public static Wd.Document CloneFrom(this Wd.Document target, string Fullname)
        {
            target.CopyStylesFromTemplate(Fullname);
            target.ApplyDocumentTheme(Fullname);

            var source = target.Application.Documents.Open(Fullname, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, false);
            try
            {
                AttachTemplate(target, source);
                CopyAllContent(target, source);
                CopyCustomDocumentProperties(target, source);
                CopyDocumentVariables(target, source);
                CopyCustomXmlParts(target, source);
                CopyTitle(target, source);
            }
            finally
            {
                source.Saved = true; // Because Word often prompts anyway..
                source.Close(Wd.WdSaveOptions.wdDoNotSaveChanges);
            }
            return target;
        }

        static void AttachTemplate(Wd.Document target, Wd.Document source)
        {
            var template = source.get_AttachedTemplate();
            object o = template.FullName;
            target.set_AttachedTemplate(ref o);
        }

        static void CopyAllContent(Wd.Document target, Wd.Document source)
        {
            source.Range().Copy();
            target.Range().Paste();
            ClipboardHelper.Clear();
        }

        static void CopyCustomDocumentProperties(Wd.Document target, Wd.Document source)
        {
            //dynamic source_props = source.CustomDocumentProperties;
            //dynamic target_props = target.CustomDocumentProperties;
            //foreach (dynamic prop in source_props)
            //    target_props.Add(prop.Name, prop.LinkToContent, prop.Type, prop.Value, prop.LinkSource);

            object source_props = source.GetType().InvokeMember("CustomDocumentProperties", BindingFlags.Default | BindingFlags.GetProperty, null, source, null);
            object target_props = target.GetType().InvokeMember("CustomDocumentProperties", BindingFlags.Default | BindingFlags.GetProperty, null, target, null);
            Type soure_typeProps = source_props.GetType();
            Type target_typeProps = target_props.GetType();
            int count = (int)soure_typeProps.InvokeMember("Count", BindingFlags.Default | BindingFlags.GetProperty, null, source_props, new object[] { });

            for (int i = 1; i <= count; i++)
            {
                object prop = soure_typeProps.InvokeMember("Item", BindingFlags.Default | BindingFlags.GetProperty, null, source_props, new object[] { i });
                Type typeProp = prop.GetType();
                var name = typeProp.InvokeMember("Name", BindingFlags.Default | BindingFlags.GetProperty, null, prop, new object[] { }).ToString();
                var value = typeProp.InvokeMember("Value", BindingFlags.Default | BindingFlags.GetProperty, null, prop, new object[] { }).ToString();

                object[] oArgs = { name, false, O.MsoDocProperties.msoPropertyTypeString, value };
                target_typeProps.InvokeMember("Add", BindingFlags.Default | BindingFlags.InvokeMethod, null, target_props, oArgs);
            }
        }

        static void CopyDocumentVariables(Wd.Document target, Wd.Document source)
        {
            foreach (Wd.Variable var in source.Variables)
                target.SetDocVar(var.Name, var.Value);
        }

        static void CopyCustomXmlParts(Wd.Document target, Wd.Document source)
        {
            foreach (O.CustomXMLPart part in source.CustomXMLParts)
                if (part.NamespaceURI != "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
                    target.CustomXMLParts.Add(part.XML);
        }

        static void CopyTitle(Wd.Document target, Wd.Document source)
        {
            target.SetBuiltInProp("Title", source.GetBuiltInProp("Title"));
        }
    }
}
