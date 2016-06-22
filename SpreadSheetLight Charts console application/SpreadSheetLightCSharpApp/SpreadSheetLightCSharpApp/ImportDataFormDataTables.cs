using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace SpreadSheetLightCSharpApp
{
    class ImportDataFormDataTables
    {
        static void Main11(string[] args)
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
                    decimal.Round((decimal)(rand.NextDouble() * 500 + 200), 4),
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

            style.FormatCode = "#,##0.0000";
            sl.SetColumnStyle(5, style);

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

            sl.SaveAs("ImportDataTable.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
