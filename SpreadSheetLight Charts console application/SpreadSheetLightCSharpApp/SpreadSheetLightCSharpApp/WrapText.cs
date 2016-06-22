using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;


namespace SpreadSheetLightCSharpApp
{
    class WrapText
    {
        static void Main7(string[] args)
        {
            SLDocument sl = new SLDocument();

            sl.SetCellValue(2, 2, @"This line should be broken into multiple lines");

            sl.SetCellValue(4, 4, "Another line that\nshould have been\nbroken into many lines");

            // this is the "trick"
            SLStyle style = sl.CreateStyle();
            style.SetWrapText(true);

            sl.SetCellStyle(2, 2, style);
            sl.SetCellStyle(4, 4, style);

            sl.SetCellValue(8, 7, "I'm tired of things\nbeing broken...");
            sl.SetCellStyle(8, 7, style);

            sl.SetColumnWidth(7, 18);
            // avoid manually setting the row height as well because by default,
            // Excel seems to take care of the "row autofitting" just fine...

            sl.SaveAs("MultipleLinesInOneCell.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
