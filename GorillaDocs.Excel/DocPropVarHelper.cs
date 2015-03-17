using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using O = Microsoft.Office.Core;
using XL = Microsoft.Office.Interop.Excel;

namespace GorillaDocs.Excel
{
    public static class DocPropVarHelper
    {
        [System.Diagnostics.DebuggerStepThrough]
        public static string GetDocProp(this XL.Workbook wb, string name)
        {
            var props = wb.CustomDocumentProperties;
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
        public static bool GetDocPropBool(this XL.Workbook wb, string name)
        {
            var props = wb.CustomDocumentProperties;
            Type typeProps = props.GetType();

            try
            {
                object prop = typeProps.InvokeMember("Item", BindingFlags.Default | BindingFlags.GetProperty, null, props, new object[] { name });
                Type typeProp = prop.GetType();
                return Convert.ToBoolean(typeProp.InvokeMember("Value", BindingFlags.Default | BindingFlags.GetProperty, null, prop, new object[] { }));
            }
            catch
            {
                // Property doesn't exist
                return false;
            }
        }
        [System.Diagnostics.DebuggerStepThrough]
        public static void SetDocProp(this XL.Workbook doc, string name, string value)
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
        public static void DeleteDocProp(this XL.Workbook wb, string name)
        {
            var props = wb.CustomDocumentProperties;
            Type typeProps = props.GetType();
            object prop = typeProps.InvokeMember("Item", BindingFlags.Default | BindingFlags.GetProperty, null, props, new object[] { name });
            typeProps.InvokeMember("Delete", BindingFlags.Default | BindingFlags.InvokeMethod, null, prop, null);
        }

        public static bool DocPropExists(this XL.Workbook wb, string name) { return !string.IsNullOrEmpty(wb.GetDocProp(name)); }
    }
}
