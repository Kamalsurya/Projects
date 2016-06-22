using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace SpreadSheetLightCSharpApp
{
    class Borders
    {
        static void Main10(string[] args)
        {
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

            sl.SaveAs("Borders.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
