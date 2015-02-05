using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;
using O = Microsoft.Office.Core;
using WF = System.Windows.Forms;

namespace GorillaDocs.Word
{
    public static class ChartHelper
    {
        public static Wd.Chart GetSelectedChartOrAddNew(this Wd.Selection selection)
        {
            if (selection.Type == Wd.WdSelectionType.wdSelectionInlineShape && selection.InlineShapes[1].Type == Wd.WdInlineShapeType.wdInlineShapeChart)
                return selection.InlineShapes[1].Chart;
            else
                return selection.InlineShapes.AddChart().Chart;
        }

        public static List<Wd.InlineShape> InlineCharts(this Wd.Document doc)
        {
            var charts = new List<Wd.InlineShape>();
            foreach (Wd.InlineShape shape in doc.InlineShapes)
                if (shape.Type == Wd.WdInlineShapeType.wdInlineShapeChart)
                    charts.Add(shape);
            return charts;
        }

        public static List<Wd.Shape> Charts(this Wd.Document doc)
        {
            var charts = new List<Wd.Shape>();
            foreach (Wd.Shape shape in doc.Shapes)
                if (shape.Type == O.MsoShapeType.msoChart)
                    charts.Add(shape);
            return charts;
        }

        public static void Update(this List<Wd.InlineShape> shapes)
        {
            foreach (Wd.InlineShape shape in shapes)
                if (shape.Type == Wd.WdInlineShapeType.wdInlineShapeChart && File.Exists(shape.LinkFormat.SourceFullName))
                    shape.LinkFormat.Update();
        }

        public static void Update(this List<Wd.Shape> shapes)
        {
            foreach (Wd.Shape shape in shapes)
                if (shape.Type == O.MsoShapeType.msoChart && File.Exists(shape.LinkFormat.SourceFullName))
                    shape.LinkFormat.Update();
        }

        public static void UpdateCharts(this Wd.Document doc)
        {
            var charts = doc.Charts();
            var inlineCharts = doc.InlineCharts();

            if ((charts.HasBrokenLinks() || inlineCharts.HasBrokenLinks()) && WF.MessageBox.Show("The document contains broken links.\n\nPress OK to fix the links.\nPress Cancel to leave them as they are.", "Fix broken links", WF.MessageBoxButtons.OKCancel, WF.MessageBoxIcon.Information) == WF.DialogResult.OK)
                doc.Application.Dialogs[Wd.WdWordDialog.wdDialogEditLinks].Show();

            charts.Update();
            inlineCharts.Update();
        }

        public static List<string> BrokenLinks(this List<Wd.Shape> shapes)
        {
            var brokenLinks = new List<string>();
            foreach (Wd.Shape shape in shapes)
                if (shape.Type == O.MsoShapeType.msoChart && !File.Exists(shape.LinkFormat.SourceFullName))
                    brokenLinks.Add(shape.LinkFormat.SourceFullName);
            return brokenLinks;
        }
        public static List<string> BrokenLinks(this List<Wd.InlineShape> shapes)
        {
            var brokenLinks = new List<string>();
            foreach (Wd.InlineShape shape in shapes)
                if (shape.Type == Wd.WdInlineShapeType.wdInlineShapeChart && !File.Exists(shape.LinkFormat.SourceFullName))
                    brokenLinks.Add(shape.LinkFormat.SourceFullName);
            return brokenLinks;
        }

        public static bool HasBrokenLinks(this List<Wd.Shape> shapes)
        {
            foreach (Wd.Shape shape in shapes)
                if (shape.Type == O.MsoShapeType.msoChart && !File.Exists(shape.LinkFormat.SourceFullName))
                    return true;
            return false;
        }
        public static bool HasBrokenLinks(this List<Wd.InlineShape> shapes)
        {
            foreach (Wd.InlineShape shape in shapes)
                if (shape.Type == Wd.WdInlineShapeType.wdInlineShapeChart && !File.Exists(shape.LinkFormat.SourceFullName))
                    return true;
            return false;
        }
    }
}
