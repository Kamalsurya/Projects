using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace SpreadSheetLightCSharpApp
{
    class WorkSheetProtection
    {
        static void Main6(string[] args)
        {
            SLDocument sl = new SLDocument();

            // Note that there's no password protection.
            // This just prevents the casual user from editing the
            // worksheet.

            SLSheetProtection sp = new SLSheetProtection();
            sp.AllowInsertRows = true;
            sp.AllowInsertColumns = false;
            sp.AllowFormatCells = true;
            sp.AllowDeleteColumns = true;
            sl.ProtectWorksheet(sp);

            // Use this to unprotect the currently selected worksheet.
            //sl.UnprotectWorksheet();

            // Note that this only unprotects worksheet without password protection.

            sl.SetCellValue(2, 2, "I'm protected. Sort of...");

            sl.SaveAs("WorksheetProtection.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
