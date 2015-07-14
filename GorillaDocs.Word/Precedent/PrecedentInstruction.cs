using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.Precedent
{
    public class PrecedentInstruction
    {
        public string Command { get; set; }
        public string Expression { get; set; }
        public string ListItems { get; set; }
        public string ExpressionObjectType { get; set; }
        public string ExpressionObjectNamespace { get; set; }

        public dynamic ExpressionData(Wd.Document doc)
        {
            Type type = Type.GetType(ExpressionObjectType);
            string xml = doc.CustomXMLParts.SelectByNamespace(ExpressionObjectNamespace)[1].XML;

            MethodInfo method = typeof(Serializer).GetMethod("DeSerializeFromString");
            MethodInfo generic = method.MakeGenericMethod(type);
            return generic.Invoke(null, new object[] { xml });
        }

        public IList<string> GetListItems(Wd.Document doc)
        {
            dynamic data = ExpressionData(doc);
            if (ListItems.Contains("."))
            {
                var bits = Regex.Split(ListItems, @"[\[(\]\.)]");
                dynamic collection = data.GetType().GetProperty(bits[0]).GetValue(data, null);
                dynamic recip = collection[bits[1].ToVal()];
                return recip.GetType().GetProperty(bits[3]).GetValue(recip, null);
            }
            else
                return data.GetType().GetProperty(ListItems).GetValue(data, null);
        }
    }
}
