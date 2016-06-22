using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
namespace SpreadSheetLightCSharpApp
{
    class NumberFormat
    {
        static void Main11(string[] args)
        {
            SLDocument sl = new SLDocument();

            sl.SetCellValue("A1", 123456789.12345);
            sl.SetCellValue(2, 1, -123456789.12345);
            sl.SetCellValue(3, 1, new DateTime(2123, 4, 15));
            sl.SetCellValue(4, 1, 12.3456);
            sl.SetCellValue(5, 1, 12.3456);
            sl.SetCellValue("A6", 123456789.12345);

            SLStyle style = sl.CreateStyle();
            style.FormatCode = "#,##0.000";
            sl.SetCellStyle("A1", style);

            style = sl.CreateStyle();
            style.FormatCode = "$#,##0.00_);[Red]($#,##0.00)";
            sl.SetCellStyle(2, 1, style);

            style = sl.CreateStyle();
            style.FormatCode = "d mmm yyyy";
            sl.SetCellStyle(3, 1, style);

            // we can just reassign like this because the only property
            // we just used was the FormatCode property

            style.FormatCode = "0.00%";
            sl.SetCellStyle("A4", style);

            // this means "number with fractional part (2 digit denominator)"
            style.FormatCode = "# ??/??";
            sl.SetCellStyle(5, 1, style);

            style.FormatCode = "0.000E+00";
            sl.SetCellStyle(6, 1, style);

            sl.SaveAs("NumberFormat.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
