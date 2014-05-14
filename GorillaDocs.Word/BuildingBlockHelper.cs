using System;
using System.Collections.Generic;
using System.Linq;
using O = Microsoft.Office.Core;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class BuildingBlockHelper
    {
        public static Wd.Range Insert(this Wd.BuildingBlock buildingblock, Wd.Range Where, object RichText, O.MsoLanguageID cultureID)
        {
            return buildingblock.Insert(Where, RichText, (Wd.WdLanguageID)cultureID);
        }
        public static Wd.Range Insert(this Wd.BuildingBlock buildingblock, Wd.Range Where, object RichText, Wd.WdLanguageID cultureID)
        {
            Wd.Range range = buildingblock.Insert(Where, RichText);
            range.LanguageID = Wd.WdLanguageID.wdEnglishUS; // Set the default for Asian languages when typing in English
            range.LanguageID = cultureID;
            return range;
        }
    }
}
