using GorillaDocs.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs
{
    public static class OfficeHelper
    {
        public static IOffice First(this List<IOffice> offices, string name)
        {
            foreach (IOffice office in offices)
                if (office.Name == name)
                    return office;
            throw new InvalidOperationException(string.Format("'{0}' is not a valid name.", name));
        }

        public static List<IOffice> Where(this List<IOffice> offices, string TemplateName)
        {
            var temp = new List<IOffice>(offices);
            foreach (IOffice office in offices)
                if (office.ExcludedTemplates.Any(x => x == TemplateName))
                    temp.RemoveAll(x => x.Name == office.Name);
            return temp;
        }

        public static IOffice LastSelected(this List<IOffice> offices, IUserSettings settings) { return offices.FirstOrDefault(x => x.Name == settings.Last_Selected_Office); }

    }
}
