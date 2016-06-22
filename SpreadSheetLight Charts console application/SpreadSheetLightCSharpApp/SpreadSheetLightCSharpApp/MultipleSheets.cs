using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using SpreadsheetLight.Charts;

namespace SpreadSheetLightCSharpApp
{
    class MultipleSheets
    {
        static void Main3(string[] args)
        {
            using (SLDocument sl = new SLDocument())
            {
                Random rand = new Random();
                DateTime dt;
                int iNumberOfIterations = 10;
                int i, iRowIndex;
                SLStyle datestyle = sl.CreateStyle();
                datestyle.FormatCode = "dd/MM/yyyy";
                SLStyle stockstyle = sl.CreateStyle();
                stockstyle.FormatCode = "0.00";
                SLChart chart;
                double fBase = 19.5;
                double fLow = 19.49;
                double fHigh = 20.51;
                double fEpsilon = 0.1;
                int iVolumeLow = 50000;
                int iVolumeHigh = 70000;

                sl.RenameWorksheet(SLDocument.DefaultFirstSheetName, "HLC");

                sl.SetCellValue("A1", "SL Fruits Private Limited");
                sl.SetCellValue("A2", "Date");
                sl.SetCellValue("B2", "High");
                sl.SetCellValue("C2", "Low");
                sl.SetCellValue("D2", "Close");

                dt = new DateTime(2013, 1, 1);
                iRowIndex = 3;
                for (i = 0; i < iNumberOfIterations; ++i)
                {
                    if (dt.DayOfWeek != DayOfWeek.Saturday && dt.DayOfWeek != DayOfWeek.Sunday)
                    {
                        sl.SetCellValue(iRowIndex, 1, dt);
                        sl.SetCellValue(iRowIndex, 2, Math.Round(fHigh + rand.NextDouble() * fEpsilon, 2));
                        sl.SetCellValue(iRowIndex, 3, Math.Round(fLow - rand.NextDouble() * fEpsilon, 2));
                        sl.SetCellValue(iRowIndex, 4, Math.Round(rand.NextDouble() + fBase, 2));
                        ++iRowIndex;
                    }
                    dt = dt.AddDays(1);
                }

                sl.SetColumnStyle(1, datestyle);
                sl.SetColumnStyle(2, 4, stockstyle);
                sl.AutoFitColumn(1);

                chart = sl.CreateChart("A2", "D10", new SLCreateChartOptions() { RowsAsDataSeries = false });
                chart.SetChartType(SLStockChartType.HighLowClose);
                chart.SetChartPosition(11, 0, 27, 6);

                // Stock charts typically don't have data on Saturdays and Sundays, because there's
                // no trading. However, the data on the worksheet is in consecutive rows, even if
                // the dates aren't consecutive. So explicitly set the primary horizontal axis
                // as a text axis. SpreadsheetLight automatically detects that the category axis
                // has date data, so by default sets it as a date axis.
                chart.PrimaryTextAxis.SetAsCategoryAxis();
                
                sl.InsertChart(chart);

                sl.AddWorksheet("OHLC");

                sl.SetCellValue("A1", "SL Fruits Private Limited");
                sl.SetCellValue("A2", "Date");
                sl.SetCellValue("B2", "Open");
                sl.SetCellValue("C2", "High");
                sl.SetCellValue("D2", "Low");
                sl.SetCellValue("E2", "Close");

                dt = new DateTime(2013, 1, 1);
                iRowIndex = 3;
                for (i = 0; i < iNumberOfIterations; ++i)
                {
                    if (dt.DayOfWeek != DayOfWeek.Saturday && dt.DayOfWeek != DayOfWeek.Sunday)
                    {
                        sl.SetCellValue(iRowIndex, 1, dt);
                        sl.SetCellValue(iRowIndex, 2, Math.Round(rand.NextDouble() + fBase, 2));
                        sl.SetCellValue(iRowIndex, 3, Math.Round(fHigh + rand.NextDouble() * fEpsilon, 2));
                        sl.SetCellValue(iRowIndex, 4, Math.Round(fLow - rand.NextDouble() * fEpsilon, 2));
                        sl.SetCellValue(iRowIndex, 5, Math.Round(rand.NextDouble() + fBase, 2));
                        ++iRowIndex;
                    }
                    dt = dt.AddDays(1);
                }

                sl.SetColumnStyle(1, datestyle);
                sl.SetColumnStyle(2, 5, stockstyle);
                sl.AutoFitColumn(1);

                chart = sl.CreateChart("A2", "E10", new SLCreateChartOptions() { RowsAsDataSeries = false });
                chart.SetChartType(SLStockChartType.OpenHighLowClose);
                chart.SetChartPosition(11, 0, 27, 6);
                sl.InsertChart(chart);

                sl.AddWorksheet("VHLC");

                sl.SetCellValue("A1", "SL Fruits Private Limited");
                sl.SetCellValue("A2", "Date");
                sl.SetCellValue("B2", "Volume");
                sl.SetCellValue("C2", "High");
                sl.SetCellValue("D2", "Low");
                sl.SetCellValue("E2", "Close");

                dt = new DateTime(2013, 1, 1);
                iRowIndex = 3;
                for (i = 0; i < iNumberOfIterations; ++i)
                {
                    if (dt.DayOfWeek != DayOfWeek.Saturday && dt.DayOfWeek != DayOfWeek.Sunday)
                    {
                        sl.SetCellValue(iRowIndex, 1, dt);
                        sl.SetCellValue(iRowIndex, 2, rand.Next(iVolumeLow, iVolumeHigh));
                        sl.SetCellValue(iRowIndex, 3, Math.Round(fHigh + rand.NextDouble() * fEpsilon, 2));
                        sl.SetCellValue(iRowIndex, 4, Math.Round(fLow - rand.NextDouble() * fEpsilon, 2));
                        sl.SetCellValue(iRowIndex, 5, Math.Round(rand.NextDouble() + fBase, 2));
                        ++iRowIndex;
                    }
                    dt = dt.AddDays(1);
                }

                sl.SetColumnStyle(1, datestyle);
                sl.SetColumnStyle(3, 5, stockstyle);
                sl.AutoFitColumn(1);

                chart = sl.CreateChart("A2", "E10", new SLCreateChartOptions() { RowsAsDataSeries = false });
                chart.SetChartType(SLStockChartType.VolumeHighLowClose);
                chart.SetChartPosition(11, 0, 27, 6);
                sl.InsertChart(chart);

                sl.AddWorksheet("VOHLC");

                sl.SetCellValue("A1", "SL Fruits Private Limited");
                sl.SetCellValue("A2", "Date");
                sl.SetCellValue("B2", "Volume");
                sl.SetCellValue("C2", "Open");
                sl.SetCellValue("D2", "High");
                sl.SetCellValue("E2", "Low");
                sl.SetCellValue("F2", "Close");

                dt = new DateTime(2013, 1, 1);
                iRowIndex = 3;
                for (i = 0; i < iNumberOfIterations; ++i)
                {
                    if (dt.DayOfWeek != DayOfWeek.Saturday && dt.DayOfWeek != DayOfWeek.Sunday)
                    {
                        sl.SetCellValue(iRowIndex, 1, dt);
                        sl.SetCellValue(iRowIndex, 2, rand.Next(iVolumeLow, iVolumeHigh));
                        sl.SetCellValue(iRowIndex, 3, Math.Round(rand.NextDouble() + fBase, 2));
                        sl.SetCellValue(iRowIndex, 4, Math.Round(fHigh + rand.NextDouble() * fEpsilon, 2));
                        sl.SetCellValue(iRowIndex, 5, Math.Round(fLow - rand.NextDouble() * fEpsilon, 2));
                        sl.SetCellValue(iRowIndex, 6, Math.Round(rand.NextDouble() + fBase, 2));
                        ++iRowIndex;
                    }
                    dt = dt.AddDays(1);
                }

                sl.SetColumnStyle(1, datestyle);
                sl.SetColumnStyle(3, 6, stockstyle);
                sl.AutoFitColumn(1);

                chart = sl.CreateChart("A2", "F10", new SLCreateChartOptions() { RowsAsDataSeries = false });
                chart.SetChartType(SLStockChartType.VolumeOpenHighLowClose);
                chart.SetChartPosition(11, 0, 27, 6);
                sl.InsertChart(chart);
                //Response.Clear();
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.AddHeader("Content-Disposition", "attachment; filename=ChartsStock.xlsx");
                //sl.SaveAs(Response.OutputStream);
                //Response.End();
                sl.SaveAs("E:\\ChartsStock.xlsx");
            }
            Console.WriteLine("End of program");
            Console.ReadLine();


        }

    }
}
