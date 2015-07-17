using GorillaDocs.libs.PostSharp;
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

        [LoudRibbonExceptionHandler]
        void doc_BuildingBlockInsert(Wd.Range Range, string Name, string Category, string BlockType, string Template)
        {
            if (Category == "Section")
            {
                Range = Range.CollapseStart();
                if (Range.Start == Range.Paragraphs[1].Range.Start)
                {
                    Range.InsertBreak_Safe(Wd.WdBreakType.wdSectionBreakNextPage);
                    Range.Sections[1].ContinueNumbering();
                }
                else
                {
                    Range.Document.Undo();
                    throw new InvalidOperationException("Ensure that the cursor is at the beginning of a new line and try again.");
                }
            }
        }
    }
}
