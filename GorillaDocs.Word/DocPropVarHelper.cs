using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using O = Microsoft.Office.Core;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class DocPropVarHelper
    {
        [System.Diagnostics.DebuggerStepThrough]
        public static string GetDocProp(this Wd.Document doc, string name)
        {
            var props = doc.CustomDocumentProperties;
            Type typeProps = props.GetType();

            try
            {
                object prop = typeProps.InvokeMember("Item", BindingFlags.Default | BindingFlags.GetProperty, null, props, new object[] { name });
                Type typeProp = prop.GetType();
                return typeProp.InvokeMember("Value", BindingFlags.Default | BindingFlags.GetProperty, null, prop, new object[] { }).ToString();
            }
            catch
            {
                // Property doesn't exist
                return string.Empty;
            }
        }
        [System.Diagnostics.DebuggerStepThrough]
        public static void SetDocProp(this Wd.Document doc, string name, string value)
        {
            var props = doc.CustomDocumentProperties;
            Type typeProps = props.GetType();
            object prop = null;
            try
            {
                prop = typeProps.InvokeMember("Item", BindingFlags.Default | BindingFlags.GetProperty, null, props, new object[] { name });
                typeProps.InvokeMember("Item", BindingFlags.Default | BindingFlags.SetProperty, null, props, new object[] { name, value });
            }
            catch
            {
                object[] oArgs = { name, false, O.MsoDocProperties.msoPropertyTypeString, value };
                typeProps.InvokeMember("Add", BindingFlags.Default | BindingFlags.InvokeMethod, null, props, oArgs);
            }
        }

        [System.Diagnostics.DebuggerStepThrough]
        public static string GetDocProp(this Wd.Template template, string name)
        {
            var props = template.CustomDocumentProperties;
            Type typeProps = props.GetType();

            try
            {
                object prop = typeProps.InvokeMember("Item", BindingFlags.Default | BindingFlags.GetProperty, null, props, new object[] { name });
                Type typeProp = prop.GetType();
                return typeProp.InvokeMember("Value", BindingFlags.Default | BindingFlags.GetProperty, null, prop, new object[] { }).ToString();
            }
            catch
            {
                // Property doesn't exist
                return string.Empty;
            }
        }
        [System.Diagnostics.DebuggerStepThrough]
        public static void SetDocProp(this Wd.Template template, string name, string value)
        {
            var props = template.CustomDocumentProperties;
            Type typeProps = props.GetType();
            object prop = null;
            try
            {
                prop = typeProps.InvokeMember("Item", BindingFlags.Default | BindingFlags.GetProperty, null, props, new object[] { name });
                typeProps.InvokeMember("Item", BindingFlags.Default | BindingFlags.SetProperty, null, props, new object[] { name, value });
            }
            catch
            {
                object[] oArgs = { name, false, O.MsoDocProperties.msoPropertyTypeString, value };
                typeProps.InvokeMember("Add", BindingFlags.Default | BindingFlags.InvokeMethod, null, props, oArgs);
            }
        }

        [System.Diagnostics.DebuggerStepThrough]
        public static string GetBuiltInProp(this Wd.Document doc, string name)
        {
            var props = doc.BuiltInDocumentProperties;
            Type typeProps = props.GetType();

            try
            {
                object prop = typeProps.InvokeMember("Item", BindingFlags.Default | BindingFlags.GetProperty, null, props, new object[] { name });
                Type typeProp = prop.GetType();
                return typeProp.InvokeMember("Value", BindingFlags.Default | BindingFlags.GetProperty, null, prop, new object[] { }).ToString();
            }
            catch
            {
                // Property doesn't exist
                return string.Empty;
            }
        }
        [System.Diagnostics.DebuggerStepThrough]
        public static void SetBuiltInProp(this Wd.Document doc, string name, string value)
        {
            var props = doc.BuiltInDocumentProperties;
            Type typeProps = props.GetType();
            object prop = null;
            try
            {
                prop = typeProps.InvokeMember("Item", BindingFlags.Default | BindingFlags.GetProperty, null, props, new object[] { name });
                typeProps.InvokeMember("Item", BindingFlags.Default | BindingFlags.SetProperty, null, props, new object[] { name, value });
            }
            catch
            {
                object[] oArgs = { name, false, O.MsoDocProperties.msoPropertyTypeString, value };
                typeProps.InvokeMember("Add", BindingFlags.Default | BindingFlags.InvokeMethod, null, props, oArgs);
            }
        }

        [System.Diagnostics.DebuggerStepThrough]
        public static void DeleteDocProp(this Wd.Document doc, string name)
        {
            var props = doc.CustomDocumentProperties;
            Type typeProps = props.GetType();
            object prop = typeProps.InvokeMember("Item", BindingFlags.Default | BindingFlags.GetProperty, null, props, new object[] { name });
            typeProps.InvokeMember("Delete", BindingFlags.Default | BindingFlags.InvokeMethod, null, prop, null);
        }

        public static bool DocPropExists(this Wd.Document doc, string name)
        {
            return !string.IsNullOrEmpty(doc.GetDocProp(name));
        }

        public static string GetDocVar(this Wd.Document doc, string name)
        {
            foreach (Wd.Variable var in doc.Variables)
                if (var.Name == name)
                    return var.Value;
            return string.Empty;
        }
        public static void SetDocVar(this Wd.Document doc, string name, string value)
        {
            foreach (Wd.Variable var in doc.Variables)
                if (var.Name == name)
                {
                    var.Value = value;
                    return;
                }
            doc.Variables.Add(name, value);
        }

    }
}
