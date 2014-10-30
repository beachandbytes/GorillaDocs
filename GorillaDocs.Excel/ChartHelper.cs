using System;
using System.Collections.Generic;
using System.Linq;
using Xl = Microsoft.Office.Interop.Excel;

namespace GorillaDocs.Excel
{
    public static class ChartHelper
    {
        public static void InsertChart(this Xl.Worksheet Sheet, string FullName)
        {
            Sheet.Shapes.AddChart2(201, Xl.XlChartType.xlColumnClustered).Select();
            Sheet.Application.ActiveChart.ApplyChartTemplate(FullName);
        }
    }
}
