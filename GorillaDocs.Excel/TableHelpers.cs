﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using XL = Microsoft.Office.Interop.Excel;

namespace GorillaDocs.Excel
{
    public static class TableHelpers
    {
        public static XL.Range Cell(this XL.ListRow row, string column)
        {
            XL.ListObject tbl = row.Parent;
            if (!tbl.ListColumns.Exists(column))
                throw new InvalidOperationException(string.Format("The column '{0}' does not exist.", column));
            return row.Range[1, tbl.ListColumns[column].Index];
        }

        public static bool Exists(this XL.ListColumns columns, string name)
        {
            foreach (XL.ListColumn column in columns)
                if (column.Name == name)
                    return true;
            return false;
        }

        public static XL.Range NamedRange(this XL.Worksheet sheet, string name)
        {
            try
            {
                return sheet.Range[name];
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException(string.Format("Unable to find Named Range '{0}'", name), ex);
            }
        }

        public static bool Exists(this XL.Sheets sheets, string name)
        {
            foreach (XL.Worksheet sheet in sheets)
                if (sheet.Name == name)
                    return true;
            return false;
        }

        public static bool Exists(this XL.ListObjects listObjects, string name)
        {
            foreach (XL.ListObject listObject in listObjects)
                if (listObject.Name == name)
                    return true;
            return false;
        }

        public static void SetCell(this XL.ListRow row, string column, string Text)
        {
            XL.ListObject tbl = row.Parent;
            if (!tbl.ListColumns.Exists(column))
                throw new InvalidOperationException(string.Format("The column '{0}' does not exist.", column));
            XL.Range range = row.Range[1, tbl.ListColumns[column].Index];
            range.Value = Text;
        }

        public static XL.ListRow FindRow(this XL.ListObject table, string column, string Text)
        {
            foreach (XL.ListRow row in table.ListRows)
                if (row.Cell(column).Text == Text)
                    return row;
            throw new InvalidOperationException(string.Format("Unable to find a Row where '{0}'='{1}' in table '{2}'", column, Text, table.Name));
        }

    }
}
