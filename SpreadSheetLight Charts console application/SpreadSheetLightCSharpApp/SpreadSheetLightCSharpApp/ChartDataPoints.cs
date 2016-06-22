using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using SpreadsheetLight.Charts;

namespace SpreadSheetLightCSharpApp
{
    class ChartDataPoints
    {
        static void Main4(string[] args)
        {
            SLDocument sl = new SLDocument();

            sl.SetCellValue("C2", "Apple");
            sl.SetCellValue("D2", "Banana");
            sl.SetCellValue("E2", "Cherry");
            sl.SetCellValue("F2", "Durian");
            sl.SetCellValue("G2", "Elderberry");
            sl.SetCellValue("B3", "North");
            sl.SetCellValue("B4", "South");
            sl.SetCellValue("B5", "East");
            sl.SetCellValue("B6", "West");

            Random rand = new Random();
            for (int i = 3; i <= 6; ++i)
            {
                for (int j = 3; j <= 7; ++j)
                {
                    sl.SetCellValue(i, j, 9000 * rand.NextDouble() + 1000);
                }
            }

            double fChartHeight = 15.0;
            double fChartWidth = 7.5;

            SLChart chart;
            SLDataPointOptions dpoptions;

            chart = sl.CreateChart("B2", "G6");
            chart.SetChartType(SLColumnChartType.ClusteredColumn);
            chart.SetChartPosition(1, 9, 1 + fChartHeight, 9 + fChartWidth);

            dpoptions = chart.CreateDataPointOptions();
            // 45 degrees, so it's top-left corner to bottom-right corner
            dpoptions.Fill.SetLinearGradient(SpreadsheetLight.Drawing.SLGradientPresetValues.Fire, 45);
            // 0 tint, 0 transparency
            dpoptions.Line.SetSolidLine(SLThemeColorIndexValues.Accent5Color, 0, 0);
            // 3.5 point
            dpoptions.Line.Width = 3.5m;
            // 3rd data series, 4th data point
            chart.SetDataPointOptions(3, 4, dpoptions);

            sl.InsertChart(chart);

            chart = sl.CreateChart("B2", "G6");
            chart.SetChartType(SLLineChartType.StackedLine);
            chart.SetChartPosition(7, 1, 7 + fChartHeight, 1 + fChartWidth);

            dpoptions = chart.CreateDataPointOptions();
            dpoptions.Marker.Symbol = DocumentFormat.OpenXml.Drawing.Charts.MarkerStyleValues.Star;
            dpoptions.Marker.Size = 10;
            // 0 tint, 0 transparency
            dpoptions.Marker.Fill.SetSolidFill(SLThemeColorIndexValues.Accent6Color, 0, 0);
            // 1st data series, 5th data point
            chart.SetDataPointOptions(1, 5, dpoptions);

            sl.InsertChart(chart);

            chart = sl.CreateChart("B2", "G6");
            chart.SetChartType(SLPieChartType.Pie);
            chart.SetChartPosition(16, 9, 16 + fChartHeight, 9 + fChartWidth);

            dpoptions = chart.CreateDataPointOptions();
            dpoptions.Explosion = 250;
            dpoptions.Fill.SetRadialGradient(SpreadsheetLight.Drawing.SLGradientPresetValues.Rainbow2, SpreadsheetLight.Drawing.SLGradientDirectionValues.CenterToTopLeftCorner);
            // it's a pie chart, so only the 1st data series is used.
            // Then we set it on the 3rd data point.
            chart.SetDataPointOptions(1, 3, dpoptions);

            sl.InsertChart(chart);

            sl.SaveAs("ChartsDataPoints.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
