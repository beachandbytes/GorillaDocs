using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace GorillaDocs
{
    public static class XmlHelpers
    {
        public static IEnumerable<XElement> ElementsAnyNS<T>(this IEnumerable<T> source, string localName) where T : XContainer
        {
            return source.Elements().Where(e => e.Name.LocalName == localName);
        }

        public static XElement FirstAnyNS<T>(this IEnumerable<T> source, string localName) where T : XContainer
        {
            return source.Elements().Where(e => e.Name.LocalName == localName).FirstOrDefault();
        }
    }
}
