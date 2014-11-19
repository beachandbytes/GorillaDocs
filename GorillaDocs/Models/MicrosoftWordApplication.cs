using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GorillaDocs.Models
{
    public class MicrosoftWordApplication : MicrosoftApplication
    {
        public string FileExtensions { get { return "*.do??"; } }
        public string FileExtensions_Regex { get { return "^\\.do..*$"; } }
    }
}
