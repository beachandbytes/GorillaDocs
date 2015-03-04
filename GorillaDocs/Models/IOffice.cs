using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs.Models
{
    public interface IOffice
    {
        string Name { get; set; }
        string LongDateFormat { get; set; }
        string CultureCode { get; set; }

        List<FileWithCategory> GetTemplates(MicrosoftApplication app);
        List<FileWithCategory> RecentFiles { get; set; }
        List<string> ExcludedTemplates { get; set; }
    }
}
