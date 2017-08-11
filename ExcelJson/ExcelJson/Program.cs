using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data.Common;
using Newtonsoft.Json;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace ConsoleJson
{
    class Program
    {
        static void Main(string[] args)
        {
            var ExcelPath = @"C:\Users\SUJITH KALARI\Documents\Visual Studio 2015\Projects\ExcelJson\ExcelJson\bin\Debug\Bookks.xlsx";
            var sheetName = "";
            var destinationPath = @"C:\Users\SUJITH KALARI\Documents\Visual Studio 2015\Projects\ExcelJson\ExcelJson\bin\Debug\E2J.json";
            var json = "";
            Dictionary<string, int> sheet = new Dictionary<string, int>();
            sheet.Add("Text", 1);
            sheet.Add("Image", 2);
            sheet.Add("Audio", 3);
            sheet.Add("Video", 4);

            Excel.Application xlApp = null;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Range range = null;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(ExcelPath);

            var connectionString = String.Format(@"
                Provider=Microsoft.ACE.OLEDB.12.0;
                Data Source={0};
                Extended Properties=""Excel 12.0 Xml;HDR=YES""", ExcelPath);
            File.AppendAllText(destinationPath, "{ AssetList :[");
            foreach (System.Collections.Generic.KeyValuePair<string, int> sheetInfo in sheet)
            {
                sheetName = sheetInfo.Key;
                int sheetNo = sheetInfo.Value;
                File.AppendAllText(destinationPath, sheetName);
                //Creating and opening a data connection to the Excel sheet 
                using (var conn = new OleDbConnection(connectionString))
                {
                    conn.Open();

                    var cmd = conn.CreateCommand();
                    cmd.CommandText = String.Format(
                        @"SELECT * FROM [{0}$]",
                        sheetName
                    );

                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(sheetNo);

                    range = xlWorkSheet.UsedRange;
                    int rowRange = range.Rows.Count;
                    int columnRange = range.Columns.Count;

                    int columnCount = 0;
                    Console.WriteLine(columnRange);
                    Console.WriteLine(rowRange);
                    using (var rdr = cmd.ExecuteReader())
                    {
                        var query =
                       (from DbDataRecord row in rdr
                        select row).Select(x =>
                        {
                            //dynamic item = new ExpandoObject();
                            Dictionary<string, object> item = new Dictionary<string, object>();
                            for (columnCount = 0; columnCount < columnRange; columnCount++)
                            {
                                item.Add(rdr.GetName(columnCount), x[columnCount]);
                            }
                            return item;
                        });

                        //Generates JSON from the LINQ query
                        json = JsonConvert.SerializeObject(query);

                        //Write the file to the destination path 
                        File.AppendAllText(destinationPath, json);
                    }

                }
                File.AppendAllText(destinationPath, "]");
            }
            File.AppendAllText(destinationPath, "}");

        }
    }
}
