using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class FontHelper
    {
        public static void ColorRGB(this Wd.Font font, int red, int green, int blue)
        {
            font.Color = (Wd.WdColor)(red + 0x100 * green + 0x10000 * blue);
        }
    }
}
