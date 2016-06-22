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
    class ChartsDataTable
    {
        static void Main1(string[] args)
        {
            SLDocument sl = new SLDocument();
            sl.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Samient");
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

            chart = sl.CreateChart("B2", "G6");
            chart.SetChartType(SLColumnChartType.ClusteredColumn);
            chart.SetChartPosition(1, 9, 1 + fChartHeight, 9 + fChartWidth);
            chart.ShowDataTable = true;
            // use the default data table settings except for this
            chart.DataTable.ShowLegendKeys = false;
            sl.InsertChart(chart);

            chart = sl.CreateChart("B2", "G6");
            chart.SetChartType(SLColumnChartType.ClusteredColumn);
            chart.SetChartPosition(7, 1, 7 + fChartHeight, 1 + fChartWidth);
            chart.ShowDataTable = true;
            chart.DataTable.ShowHorizontalBorder = false;
            chart.DataTable.ShowVerticalBorder = false;
            chart.DataTable.ShowOutlineBorder = false;
            chart.DataTable.ShowLegendKeys = true;
            sl.InsertChart(chart);

            chart = sl.CreateChart("B2", "G6");
            chart.SetChartType(SLColumnChartType.ClusteredColumn);
            chart.SetChartPosition(16, 9, 16 + fChartHeight, 9 + fChartWidth);
            chart.ShowDataTable = true;
            // 45 degrees meaning from top-left corner to bottom-right corner
            chart.DataTable.Fill.SetLinearGradient(SpreadsheetLight.Drawing.SLGradientPresetValues.Brass, 45);
            // 0 tint, 0 transparency
            chart.DataTable.Border.SetSolidLine(SLThemeColorIndexValues.Accent5Color, 0, 0);
            // 1.2 pt
            chart.DataTable.Border.Width = 1.2m;
            chart.DataTable.Shadow.SetPreset(SpreadsheetLight.Drawing.SLShadowPresetValues.OuterDiagonalTopRight);

            SLFont font = sl.CreateFont();
            // 18 point
            font.SetFont(FontSchemeValues.Major, 18);
            // 80% lightening
            font.SetFontThemeColor(SLThemeColorIndexValues.Accent5Color, 0.8);
            chart.DataTable.SetFont(font);

            sl.InsertChart(chart);

            sl.SaveAs("E:\\ChartsDataTable.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
