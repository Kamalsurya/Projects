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
    class ChartDataLabels
    {
        static void Main13(string[] args)
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

            // use SLGroupDataLabelOptions for an entire data series
            SLGroupDataLabelOptions gdloptions;
            // use SLDataLabelOptions for a specific data label in a specific data series
            SLDataLabelOptions dloptions;

            chart = sl.CreateChart("B2", "G6");
            chart.SetChartType(SLColumnChartType.ClusteredColumn);
            chart.SetChartPosition(1, 9, 1 + fChartHeight, 9 + fChartWidth);

            gdloptions = chart.CreateGroupDataLabelOptions();
            gdloptions.ShowValue = true;
            // 2nd data series
            chart.SetGroupDataLabelOptions(2, gdloptions);

            sl.InsertChart(chart);

            chart = sl.CreateChart("B2", "G6");
            chart.SetChartType(SLLineChartType.StackedLine);
            chart.SetChartPosition(7, 1, 7 + fChartHeight, 1 + fChartWidth);

            gdloptions = chart.CreateGroupDataLabelOptions();
            gdloptions.ShowValue = true;
            gdloptions.FormatCode = "0.00";
            // set to false so the data don't link to the source data for the number format
            gdloptions.SourceLinked = false;
            // 4th data series
            chart.SetGroupDataLabelOptions(4, gdloptions);

            dloptions = chart.CreateDataLabelOptions();
            dloptions.ShowSeriesName = true;
            dloptions.ShowCategoryName = true;
            dloptions.ShowValue = true;
            dloptions.ShowLegendKey = true;
            dloptions.FormatCode = "0.00";
            // set to false so the data don't link to the source data for the number format
            dloptions.SourceLinked = false;
            dloptions.SetLabelPosition(DocumentFormat.OpenXml.Drawing.Charts.DataLabelPositionValues.Right);
            // 2nd data series, 4th data point
            chart.SetDataLabelOptions(2, 4, dloptions);

            sl.InsertChart(chart);

            chart = sl.CreateChart("B2", "G6");
            chart.SetChartType(SLPieChartType.Pie);
            chart.SetChartPosition(16, 9, 16 + fChartHeight, 9 + fChartWidth);

            dloptions = chart.CreateDataLabelOptions();
            // you can set a custom label text
            dloptions.SetLabelText("Have Pie, Eat Pie");
            // 0 degrees, so it's from left to right
            dloptions.Fill.SetLinearGradient(SpreadsheetLight.Drawing.SLGradientPresetValues.Silver, 0);
            // 0 tint, 0 transparency
            dloptions.Border.SetSolidLine(SLThemeColorIndexValues.Accent6Color, 0, 0);
            // it's a pie chart, so the only useful data series is the 1st one.
            // Then we apply it to the 3rd data point
            chart.SetDataLabelOptions(1, 3, dloptions);

            sl.InsertChart(chart);

            sl.SaveAs("ChartsDataLabels.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
