using GorillaDocs.libs.PostSharp;
using GorillaDocs.ViewModels;
using GorillaDocs.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    [Log]
    public static class TableHelper
    {
        public static void InsertTableHorizontal(this Wd.Range range)
        {
            InsertTable(range, FormatHorizontalTable);
        }
        public static void InsertTableHorizontalShaded(this Wd.Range range)
        {
            InsertTable(range, FormatHorizontalTable, true);
        }
        public static void InsertTableVertical(this Wd.Range range)
        {
            InsertTable(range, FormatVerticalTable);
        }
        public static void InsertTableVerticalShaded(this Wd.Range range)
        {
            InsertTable(range, FormatVerticalTable, true);
        }
        static void FormatHorizontalTable(Wd.Table table, bool Shaded)
        {
            Wd.Styles styles = table.Range.Document.Styles;
            string style = "Table Horizontal Blue";
            if (Shaded)
                style = "Table Horizontal Shaded Blue";
            if (styles.Exists(style))
                table.set_Style(style);
            if (Shaded)
            {
                if (styles.Exists("Table Heading White"))
                    table.Rows[1].Range.SetStyle("Table Heading White");
            }
            else
            {
                if (styles.Exists("Table Heading"))
                    table.Rows[1].Range.SetStyle("Table Heading");
            }
            if (styles.Exists("Table Text"))
                for (int i = 2; i <= table.Rows.Count; i++)
                    table.Rows[i].Range.SetStyle("Table Text");
        }
        static void FormatVerticalTable(Wd.Table table, bool Shaded)
        {
            Wd.Styles styles = table.Range.Document.Styles;
            string style = "Table Vertical Blue";
            if (Shaded)
                style = "Table Vertical Shaded Blue";
            if (styles.Exists(style))
                table.set_Style(style);
            if (Shaded)
            {
                if (styles.Exists("Table Heading White"))
                    foreach (Wd.Cell cell in table.Columns[1].Cells)
                        cell.Range.SetStyle("Table Heading White");
            }
            else
            {
                if (styles.Exists("Table Heading"))
                    foreach (Wd.Cell cell in table.Columns[1].Cells)
                        cell.Range.SetStyle("Table Heading");
            }
            if (styles.Exists("Table Text"))
                for (int i = 2; i <= table.Columns.Count; i++)
                    foreach (Wd.Cell cell in table.Columns[i].Cells)
                        cell.Range.SetStyle("Table Text");
        }

        delegate void FormatTable(Wd.Table table, bool Shaded);
        static void InsertTable(Wd.Range range, FormatTable FormatTable, bool Shaded = false)
        {
            var viewModel = new AddTableViewModel();
            var view = new AddTableView(viewModel);
            view.ShowDialog();
            if (view.DialogResult == true)
            {
                AddTableHeading(ref range, viewModel.TableHeading);
                Wd.Table table = range.Tables.Add(range, viewModel.NumberOfRows, viewModel.NumberOfColumns, Type.Missing, Wd.WdAutoFitBehavior.wdAutoFitWindow);
                FormatTable(table, Shaded);
                AddTableSource(table, viewModel.TableSource);
            }
        }
        static void AddTableHeading(ref Wd.Range range, string tableHeading)
        {
            if (!string.IsNullOrEmpty(tableHeading))
            {
                range.InsertParagraphBefore();
                range.Text = tableHeading + "\n";
                if (range.Document.Styles.Exists("Table Heading - External"))
                    range.SetStyle("Table Heading - External");
                range.InsertParagraphAfter();
                range = range.Characters.Last;
            }
        }
        static void AddTableSource(Wd.Table table, string tableSource)
        {
            if (!string.IsNullOrEmpty(tableSource))
            {
                Wd.Range range = table.Range;
                range.CollapseEnd().InsertParagraphBefore();
                range.Text = tableSource + "\n";
                if (range.Document.Styles.Exists("Table Source"))
                    range.SetStyle("Table Source");
            }
        }

        public static bool ContainsTableCell(this Wd.Range range)
        {
            return string.IsNullOrEmpty(range.Text) ? false : range.Text.Contains("\r\a");
        }

        public static void SetAutoFit(this Wd.Tables tables)
        {
            foreach (Wd.Table table in tables)
            {
                table.AutoFitBehavior(Wd.WdAutoFitBehavior.wdAutoFitWindow);
                table.AllowAutoFit = false;
            }
        }
    }
}
