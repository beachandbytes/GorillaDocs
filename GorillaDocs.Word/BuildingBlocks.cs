using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public class BuildingBlocks
    {
        readonly Wd.Document doc;
        public BuildingBlocks(Wd.Document doc)
        {
            this.doc = doc;
            this.doc.BuildingBlockInsert += doc_BuildingBlockInsert;
        }

        void doc_BuildingBlockInsert(Wd.Range Range, string Name, string Category, string BlockType, string Template)
        {
            if (Category == "Section")
            {
                Range.CollapseStart().InsertBreak(Wd.WdBreakType.wdSectionBreakNextPage);
                Range.Sections[1].ContinueNumbering();
            }
        }
    }
}
