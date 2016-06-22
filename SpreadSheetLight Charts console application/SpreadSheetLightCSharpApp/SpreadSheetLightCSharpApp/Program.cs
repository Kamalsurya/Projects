using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using SpreadsheetLight.Charts;
using System.Data;
using System.IO;
namespace SpreadSheetLightCSharpApp
{
    class Program
    {
        static void Main(string[] args)
        {

            Chartclmdso();
            return;

            SLDocument sl = new SLDocument();

            sl.SetCellValue("E5", "Prison");

            SLStyle style = sl.CreateStyle();
            style.Border.LeftBorder.BorderStyle = BorderStyleValues.Thick;
            style.Border.LeftBorder.Color = System.Drawing.Color.BlanchedAlmond;

            style.Border.BottomBorder.BorderStyle = BorderStyleValues.DashDotDot;
            style.Border.BottomBorder.Color = System.Drawing.Color.Brown;

            style.SetRightBorder(BorderStyleValues.Hair, System.Drawing.Color.Blue);
            // Alternatively, use the "long-form" version:
            // style.Border.RightBorder.BorderStyle = BorderStyleValues.Hair;
            // style.Border.RightBorder.Color = System.Drawing.Color.Blue;

            style.SetTopBorder(BorderStyleValues.Double, SLThemeColorIndexValues.Accent6Color);
            // Alternatively, use the "long-form" version:
            // style.Border.TopBorder.BorderStyle = BorderStyleValues.Double;
            // style.Border.TopBorder.SetBorderThemeColor(SLThemeColorIndexValues.Accent6Color);

            // The "0.2" means "lightens the accent 3 colour by 20%".
            // A negative value darkens the given theme colour.
            style.SetDiagonalBorder(BorderStyleValues.MediumDashDotDot, SLThemeColorIndexValues.Accent3Color, 0.2);
            // Alternatively, use the "long-form" version:
            // style.Border.DiagonalBorder.BorderStyle = BorderStyleValues.MediumDashDotDot;
            // style.Border.DiagonalBorder.SetBorderThemeColor(SLThemeColorIndexValues.Accent3Color, 0.2);

            style.Border.DiagonalUp = true;
            style.Border.DiagonalDown = true;
            sl.SetCellStyle(5, 5, style);

            sl.SaveAs("E:\\Borders.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }

        private static void TestDS()
        {
            Random rand = new Random();

            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("Product", typeof(string));
            dt.Columns.Add("IP Address", typeof(string));
            dt.Columns.Add("Date (UTC)", typeof(DateTime));
            dt.Columns.Add("Size (MB)", typeof(double));
            dt.Columns.Add("Cost", typeof(decimal));

            for (int i = 0; i < 10; ++i)
            {
                dt.Rows.Add(string.Format("Prod{0}", rand.Next(5)),
                    string.Format("{0}.{1}.{2}.{3}", rand.Next(256), rand.Next(256), rand.Next(256), rand.Next(256)),
                    DateTime.UtcNow.AddDays(rand.NextDouble() * 20),
                  500 + 200,
                    decimal.Round((decimal)(rand.NextDouble() * 20 + 5), 2));
            }

            SLDocument sl = new SLDocument();

            int iStartRowIndex = 3;
            int iStartColumnIndex = 2;

            sl.ImportDataTable(iStartRowIndex, iStartColumnIndex, dt, true);

            // This part sets the style, but you might be using a template file,
            // so the styles are probably already set.

            SLStyle style = sl.CreateStyle();
            style.FormatCode = "yyyy/mm/dd hh:mm:ss";
            sl.SetColumnStyle(4, style);

            style.FormatCode = "$#,##0.00";
            sl.SetColumnStyle(6, style);

            // The next part is optional, but it shows how you can set a table on your
            // data based on your DataTable's dimensions.

            // + 1 because the header row is included
            // - 1 because it's a counting thing, because the start row is counted.
            int iEndRowIndex = iStartRowIndex + dt.Rows.Count + 1 - 1;
            // - 1 because it's a counting thing, because the start column is counted.
            int iEndColumnIndex = iStartColumnIndex + dt.Columns.Count - 1;
            SLTable table = sl.CreateTable(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
            table.SetTableStyle(SLTableStyleTypeValues.Medium17);
            table.HasTotalRow = true;
            table.SetTotalRowFunction(5, SLTotalsRowFunctionValues.Sum);
            sl.InsertTable(table);

            sl.SaveAs("E:/ChartsColumnDataSeriesOptions.xlsx");
        }

        private static void TestTbl()
        {
            SLDocument sl = new SLDocument();

            int i, j;
            for (i = 2; i <= 12; ++i)
            {
                for (j = 2; j <= 6; ++j)
                {
                    if (i == 2)
                    {
                        sl.SetCellValue(i, j, string.Format("Col{0}", j));
                    }
                    else
                    {
                        sl.SetCellValue(i, j, i * j);
                    }
                }
            }

            // tabular data ranges from B2:F12, inclusive of a header row
            SLTable tbl = sl.CreateTable("B2", "F12");
            tbl.SetTableStyle(SLTableStyleTypeValues.Medium9);

            sl.InsertTable(tbl);

            sl.SaveAs("E:/Tables.xlsx");

        }

        private static void RestCoptRC()
        {
            SLDocument sl = new SLDocument();

            for (int i = 1; i < 20; ++i)
            {
                for (int j = 1; j < 15; ++j)
                {
                    sl.SetCellValue(i, j, string.Format("R{0}C{1}", i, j));
                }
            }

            // copy row 3 to row 21
            sl.CopyRow(3, 21);

            // cuts rows 15 through 18 and paste it to row 23
            // So row 15 is at row 23, row 16 is at row 24 and so on.
            // Default behaviour is to copy-and-paste (false).
            sl.CopyRow(15, 18, 23, true);

            // cuts column 10 and paste it to column 5
            sl.CopyColumn(10, 5, true);

            // copy rows 10 through 12 to column 11
            // So column 10 is at column 11, column 11 is at column 12 and so on.
            // Note that column 10 is currently blank because of the cut-and-paste
            // operation above. So column 11 is also blank.
            sl.CopyColumn(10, 12, 11);

            // You can also copy row and column styles.

            SLStyle rowstyle = sl.CreateStyle();
            rowstyle.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Accent1Color, SLThemeColorIndexValues.Accent2Color);

            SLStyle colstyle = sl.CreateStyle();
            colstyle.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Accent5Color, SLThemeColorIndexValues.Accent6Color);

            // set rows 3 through 10 with the given style
            sl.SetRowStyle(3, 10, rowstyle);
            // set columns 5 through 8 with the given style
            sl.SetColumnStyle(5, 8, colstyle);

            // copy the style from row 5 to rows 12 through 15
            sl.CopyRowStyle(5, 12, 15);
            // copy the style from column 7 to column 2
            sl.CopyColumnStyle(7, 2);

            sl.SaveAs("E:/CopyRowColumn.xlsx");

        }

        private static void Chartclmdso()
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

            SLChart chart;

            chart = sl.CreateChart("B2", "G6");
            chart.SetChartType(SLBarChartType.ClusteredBar);
            chart.SetChartPosition(7, 1, 22, 8.5);

            SLDataSeriesOptions dso;
            // get the options from the 2nd data series
            dso = chart.GetDataSeriesOptions(2);
            // 10% transparency
            dso.Fill.SetSolidFill(System.Drawing.Color.Yellow, 0);
            // Set on the 2nd data series.
            // Make sure you set the options on the correct data series index that
            // you got it from.
            // Or not, depending on what you want to achieve...
            chart.SetDataSeriesOptions(2, dso);

            chart.Legend.LegendPosition = DocumentFormat.OpenXml.Drawing.Charts.LegendPositionValues.TopRight;
            chart.Legend.Fill.SetRadialGradient(SpreadsheetLight.Drawing.SLGradientPresetValues.Gold, SpreadsheetLight.Drawing.SLGradientDirectionValues.Center);
            // 0% transparency
            chart.Legend.Border.SetSolidLine(System.Drawing.Color.Orange, 0);
            chart.Legend.Shadow.SetPreset(SpreadsheetLight.Drawing.SLShadowPresetValues.PerspectiveDiagonalUpperLeft);

            sl.InsertChart(chart);
            sl.SaveAs("E:/ChartsColumnDataSeriesOptions.xlsx");
        }

        private void ChartDataSeriesOptns()
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

            SLChart chart = sl.CreateChart("B2", "G6");
            chart.SetChartType(SLColumnChartType.ClusteredColumn3D);
            chart.SetChartPosition(7, 1, 22, 8.5);

            SLDataSeriesOptions dso;
            // get the options from the 2nd data series
            dso = chart.GetDataSeriesOptions(2);
            dso.Shape = DocumentFormat.OpenXml.Drawing.Charts.ShapeValues.Pyramid;
            // 10% transparency
            dso.Fill.SetSolidFill(System.Drawing.Color.MediumOrchid, 10);
            // Set on the 2nd data series.
            // Make sure you set the options on the correct data series index that
            // you got it from.
            // Or not, depending on what you want to achieve...
            chart.SetDataSeriesOptions(2, dso);

            dso = chart.GetDataSeriesOptions(4);
            // in this case, the shadow is almost imperceptible. Just look harder, ok?
            dso.Shadow.SetPreset(SpreadsheetLight.Drawing.SLShadowPresetValues.PerspectiveDiagonalUpperRight);
            // 0% transparency
            dso.Line.SetSolidLine(System.Drawing.Color.Orange, 0);
            chart.SetDataSeriesOptions(4, dso);

            sl.InsertChart(chart);
            sl.SaveAs("E:/ChartsColumnDataSeriesOptions.xlsx");
        }
    }
}

