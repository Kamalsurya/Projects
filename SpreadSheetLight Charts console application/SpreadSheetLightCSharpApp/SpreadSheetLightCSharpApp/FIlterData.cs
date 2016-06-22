using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace SpreadSheetLightCSharpApp
{
    class FIlterData
    {
        static void Main5(string[] args)
        {
            SLDocument sl = new SLDocument();

            sl.SetCellValue(2, 2, "I");
            sl.SetCellValue(2, 3, "Love");
            sl.SetCellValue(2, 4, "These");
            sl.SetCellValue(2, 5, "Filtration");
            sl.SetCellValue(2, 6, "Tablets");

            sl.SetCellValue(3, 2, 1);
            sl.SetCellValue(3, 3, 2);
            sl.SetCellValue(3, 4, 3);
            sl.SetCellValue(3, 5, 4);
            sl.SetCellValue(3, 6, 5);

            sl.SetCellValue(4, 2, 6);
            sl.SetCellValue(4, 3, 7);
            sl.SetCellValue(4, 4, 8);
            sl.SetCellValue(4, 5, 9);
            sl.SetCellValue(4, 6, 10);

            sl.Filter("B2", "F4");

            sl.SaveAs("FilterData.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
