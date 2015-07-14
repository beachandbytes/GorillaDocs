using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class ColourHelper
    {
        public static void BackgroundPatternColorRGB(this Wd.Shading shading, int red, int green, int blue)
        {
            shading.BackgroundPatternColor = (Wd.WdColor)(red + 0x100 * green + 0x10000 * blue);
        }
    }
}
