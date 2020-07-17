using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Support.UI;

namespace ConsoleApp1
{
    class Program
    {
        public static string sql = @"Data Source = DESKTOP-8GEUI6U; Initial Catalog = GameAndApp; Integrated Security = True";

        public static List<GameAndApp> apps = new List<GameAndApp>();

        public static void ReadExcel()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\ADMIN\Desktop\\ahihi.csv");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

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
                AddToDatabase(gameAndApps, i);
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

        public static void AddToDatabase(List<GameAndApp> games, int i)
        {
            SqlConnection connection = new SqlConnection(sql);
            connection.Open();
            string cmdText = "INSERT INTO GAMEIOS_TABLE(Name, URLApp, AGEApp, Pricing, Size, Category, Description, Developer) " +
                "VALUES ('" + games[0].name + "', '" + games[0].urlApp + "', '" + games[0].ageApp + "', '" + games[0].pricing + "', '" + games[0].size + "', '" + games[0].category + "', '" + games[0].description + "', '" + games[0].developer + "')";
            SqlCommand sqlCommand = new SqlCommand(cmdText, connection);
            sqlCommand.ExecuteNonQuery();
            connection.Close();
            Console.WriteLine("Đã xong " + (i - 1));

        }

        public static void ReadData()
        {
            SqlConnection connection = new SqlConnection(sql);
            connection.Open();
            SqlCommand sqlCommand = connection.CreateCommand();
            sqlCommand.CommandText = "Select * from GAMEIOS_TABLE";
            SqlDataReader reader = sqlCommand.ExecuteReader();
            while (reader.Read())
            {
                string description = "z" + reader[7].ToString() + "z";
                description = description.Replace("zDescription", "").Replace("morez", "");
                apps.Add(new GameAndApp(reader[1].ToString(), reader[2].ToString(), reader[3].ToString(), reader[4].ToString(), reader[5].ToString(), reader[6].ToString(), description, reader[8].ToString() ));
                
            }
        }

        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            //ReadExcel();
            ReadData();


            Console.Write("Gõ run để chạy: ");
            string run = Console.ReadLine();
            if (run == "run")
            {
                ChromeDriver chromeDriver = new ChromeDriver();
                chromeDriver.Url = "https://ahihisoftware.com/login";
                chromeDriver.Navigate();

                var email = chromeDriver.FindElementByXPath("/html/body/div/div[2]/form/div[1]/input");
                email.SendKeys("nguyencuonggostack@gmail.com");

                var password = chromeDriver.FindElementByXPath("/html/body/div/div[2]/form/div[2]/input");
                password.SendKeys("Toang123!@#");

                var login = chromeDriver.FindElementByXPath("/html/body/div/div[2]/form/div[3]/div[2]/button");
                login.Click();

                var createApp = chromeDriver.FindElementByXPath("/html/body/div/aside/section/ul/li[2]/ul/li[2]/a");
                createApp.Click();

                for (int i = 0; i < apps.Count; i++)
                {
                    var title = chromeDriver.FindElementByXPath("/html/body/div/div/section[2]/div/div/div/form/div[1]/div[1]/input");
                    title.SendKeys(apps[1].name);

                    var description = chromeDriver.FindElementByXPath("/html/body/div/div/section[2]/div/div/div/form/div[1]/div[2]/input");
                    description.SendKeys(apps[1].name);

                    var category = chromeDriver.FindElementByName("category");
                    var selectCategory = new SelectElement(category);
                    selectCategory.SelectByIndex(9);

                    var platform = chromeDriver.FindElementByName("platform");
                    var selectPlatform = new SelectElement(platform);
                    selectPlatform.SelectByIndex(2);

                    var type = chromeDriver.FindElementByName("type");
                    var selectType = new SelectElement(type);
                    selectType.SelectByIndex(2);


                    var developer = chromeDriver.FindElementByXPath("/html/body/div/div/section[2]/div/div/div/form/div[1]/div[4]/div[1]/div/input");
                    developer.SendKeys(apps[1].developer);

                    var urlApp = chromeDriver.FindElementByXPath("/html/body/div/div/section[2]/div/div/div/form/div[1]/div[4]/div[2]/div/input");
                    urlApp.SendKeys(apps[1].urlApp);

                    var fileSize = chromeDriver.FindElementByXPath("/html/body/div/div/section[2]/div/div/div/form/div[1]/div[5]/div[2]/div/input");
                    fileSize.SendKeys(apps[1].size);

                    var fullDescription = chromeDriver.FindElementByClassName("wysihtml5-sandbox");
                    
                    fullDescription.SendKeys(apps[1].description.Trim());
                    Console.WriteLine("asdasd");
                }


            }
            Console.ReadKey();
        }
    }
}
