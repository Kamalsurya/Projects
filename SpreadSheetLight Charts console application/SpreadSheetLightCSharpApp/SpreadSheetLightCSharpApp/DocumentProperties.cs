using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace SpreadSheetLightCSharpApp
{
    class DocumentProperties
    {
        static void Main9(string[] args)
        {
            SLDocument sl = new SLDocument();

            sl.DocumentProperties.Creator = "Kylie";
            sl.DocumentProperties.ContentStatus = "Secret";
            sl.DocumentProperties.Title = "Bali tryst with hubby";
            sl.DocumentProperties.Description = "Secret trip plan to Bali with Randy";

            // the rest of this is extra. But you know, you could continue reading...

            sl.RenameWorksheet(SLDocument.DefaultFirstSheetName, "BaliTrip");
            sl.SetCellValue(2, 2, "Note to self: insert appropriate picture of seaside here");

            sl.AddWorksheet("Secret");
            sl.SetCellValue(2, 2, "Randy, if you're reading this, STOP!");
            sl.SetCellValue(3, 2, "Close this immediately!");
            sl.SetCellValue(5, 2, "I mean it!");

            sl.SetCellValue(17, 1, "Notes");
            sl.SetCellValue(18, 1, "Where's that lipstick with that color Randy likes...");

            sl.AddWorksheet("SuperSecret");
            sl.SetCellValue(1, 2, "RANDY!! What did I tell you?? STOP READING!");

            SLStyle style = sl.CreateStyle();
            style.SetFont("Impact", 24);
            style.Font.Underline = UnderlineValues.Single;
            sl.SetCellStyle(1, 2, style);

            sl.SetCellValue(5, 2, "Tasha was telling me this technique that's gonna make him... *cough*");
            sl.SetCellValue(6, 2, "She told me to just wear, uhm, just the lipstick...");

            sl.SetCellValue(10, 2, "I wonder if it'll work...");

            sl.SetCellValue(15, 2, "RANDY!!");
            sl.SetCellStyle(15, 2, style);

            sl.SaveAs("DocumentProperties.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
