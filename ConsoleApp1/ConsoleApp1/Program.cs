using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp1
{
    class Program
    {
        public string sql = @"Data Source = DESKTOP-8GEUI6U; Initial Catalog = PrintfulTshirt; Integrated Security = True";
        public static void ReadExcel()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\ADMIN\Desktop\\ahihi.csv");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 2; i <= rowCount; i++)
            {
                string name = "null";
                string urlApp = "null";
                string ageApp = "null";
                string pricing = "null";
                string size = "null";
                string category = "null";
                string description = "null";
                string developer = "null";


                if (xlRange.Cells[i, 1].Value2 != null) { name = xlRange.Cells[i, 3].Value2.ToString().Replace("'", @"''"); }
                if (xlRange.Cells[i, 2].Value2 != null) { urlApp = xlRange.Cells[i, 4].Value2.ToString().Replace("'", @"''"); }
                if (xlRange.Cells[i, 3].Value2 != null) { ageApp = xlRange.Cells[i, 6].Value2.ToString().Replace("'", @"''"); }
                if (xlRange.Cells[i, 4].Value2 != null) { pricing = xlRange.Cells[i, 10].Value2.ToString().Replace("'", @"''"); }
                if (xlRange.Cells[i, 5].Value2 != null) { size = xlRange.Cells[i, 12].Value2.ToString().Replace("'", @"''"); }
                if (xlRange.Cells[i, 6].Value2 != null) { category = xlRange.Cells[i, 13].Value2.ToString().Replace("'", @"''"); }
                if (xlRange.Cells[i, 7].Value2 != null) { description = xlRange.Cells[i, 15].Value2.ToString().Replace("'", @"''"); }
                if (xlRange.Cells[i, 7].Value2 != null) { developer = xlRange.Cells[i, 7].Value2.ToString().Replace("'", @"''"); }

                List<GameAndApp> gameAndApps = new List<GameAndApp>();
                gameAndApps.Add(new GameAndApp(name, urlApp, ageApp, pricing, size, category, description, developer));

            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        public static void AddToDatabase(List<GameAndApp> games)
        {

        }

        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            ReadExcel();

            Console.ReadKey();
        }
    }
}
