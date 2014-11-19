using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GorillaDocs.Models
{
    public interface MicrosoftApplication
    {
        string FileExtensions { get; }
        string FileExtensions_Regex { get; }
    }
}
