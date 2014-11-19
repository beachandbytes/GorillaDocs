using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GorillaDocs.Models
{
    public class MicrosoftPowerPointApplication : MicrosoftApplication
    {
        public string FileExtensions { get { return "*.p?t?"; } }
        public string FileExtensions_Regex { get { return "^\\.p.t.*$"; } }
    }
}
