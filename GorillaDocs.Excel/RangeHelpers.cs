using System;
using System.Collections.Generic;
using System.Linq;
using XL = Microsoft.Office.Interop.Excel;

namespace GorillaDocs.Excel
{
    public static class RangeHelpers
    {
        public static bool Exists(this XL.Names names, string NamedRange)
        {
            foreach (XL.Name name in names)
                if (name.Name == NamedRange)
                    return true;
            return false;
        }

        public static XL.Range NamedRange(this XL.Workbook wb, string value)
        {
            XL.Name name = wb.Names.Item(value);
            return name.RefersToRange;
        }

        public static bool AllCellsHaveValidation(this XL.Range range)
        {
            try
            {
                foreach (XL.Range cell in range.Cells)
                {
                    var x = cell.Validation.Type;
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static void SetNamedRange(this XL.Range range, string name)
        {
            XL.Workbook wb = range.Worksheet.Parent;
            wb.Names.Add(name, range);
        }

        public static XL.Range FirstOrDefaultRange(this XL.Names names, string Value)
        {
            foreach (XL.Name name in names)
                if (name.Name == Value)
                    return name.RefersToRange;
            return null;
        }
    }
}
