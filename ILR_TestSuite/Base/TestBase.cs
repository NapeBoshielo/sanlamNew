using NUnit.Framework;

using OpenQA.Selenium;

using OpenQA.Selenium.Chrome;

using System.Threading;

using OpenQA.Selenium.Support.UI;

using System.IO;

using System;

using OpenQA.Selenium.Firefox;
using System.Data.OleDb;
using System.Data;

namespace ILR_TestSuite

{

    [TestFixture]

    public class TestBase

    {

        private ChromeOptions _chromeOptions;

        public IWebDriver _driver;

        private string _userName;

        private string _password;

        public string _screenShotFolder;

        public string _connString;



        [SetUp]

        public void StartBrowser()



        {

            _driver = new ChromeDriver("C:/Code/bin");


            _screenShotFolder = $@"C:\Code​{ScreenShotDailyFolderName()}​\";

            new DirectoryInfo(_screenShotFolder).Create();

            
        }

        private static string ScreenShotDailyFolderName()

        {

            return DateTime.Now.ToString("yyyyMMdd").Replace("AM", string.Empty).Replace("PM", string.Empty);

        }

        public void TakeScreenshot(IWebDriver driver, string filePath, string fileName)

        {

            if (!Directory.Exists(filePath))

                new DirectoryInfo(filePath).Create();







            ITakesScreenshot ssdriver = driver as ITakesScreenshot;

            Screenshot screenshot = ssdriver.GetScreenshot();

            fileName = $"{fileName}{ScreenShotTime()}.png";

            screenshot.SaveAsFile($"{filePath}{fileName}", ScreenshotImageFormat.Png);

            // byte[] byteArray = ((ITakesScreenshot)driver).GetScreenshot().AsByteArray;

            //Bitmap ss  = new Bitmap(new System.IO.MemoryStream(byteArray));

            // screenshot.SaveAsFile(String.Format(fileName, ImageFormat.Png));









        }

        private static string ScreenShotTime()

        {

            return DateTime.Now.TimeOfDay.ToString().Replace(":", "_").Replace(".", "_");

        }



        public IWebDriver SiteConnection()

        {
         

            _driver.Url = "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?";

            _userName = "G992107";//TODO add your user name and password

            _password = "G992107/d";

            _driver.Manage().Window.Maximize();

            System.Threading.Thread.Sleep(2000);

            _driver.Manage().Window.Maximize();

            System.Threading.Thread.Sleep(2000);

            IWebElement loginTextBox = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[2]/input"));
            IWebElement passwordTextBox = _driver.FindElement(By.CssSelector("#CntContentsDiv2 > table > tbody > tr:nth-child(3) > td:nth-child(2) > input[type=password]"));
            IWebElement loginBtn = _driver.FindElement(By.CssSelector("#GBLbl-1 > span > a"));
            loginTextBox.SendKeys(_userName);
            System.Threading.Thread.Sleep(2000);
            passwordTextBox.SendKeys(_password);
            System.Threading.Thread.Sleep(2000);
            loginBtn.Click();
            System.Threading.Thread.Sleep(2000);
            return _driver;
        }


        public string GetPolicyNoFromExcell(string sheet, string function)
        {

            _connString = "Provider= Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:/Users/G992107/Documents/GitHub/ILR_TestSuite/ILR_TestSuite/MIP UAT Test Scenarios/Rate_Table_Safrican Just Funeral.xlsx" + ";Extended Properties='Excel 8.0;HDR=Yes'";

            string conRef = "";
            using (OleDbConnection conn = new OleDbConnection(_connString))
            {
                try
                {

                    // Open connection
                    conn.Open();
                    string cmdQuery = "SELECT * FROM [" + sheet + "$]";

                    OleDbCommand cmd = new OleDbCommand(cmdQuery, conn);

                    // Create new OleDbDataAdapter
                    OleDbDataAdapter oleda = new OleDbDataAdapter();

                    oleda.SelectCommand = cmd;

                    // Create a DataSet which will hold the data extracted from the worksheet.
                    DataSet ds = new DataSet();

                    // Fill the DataSet from the data extracted from the worksheet.
                    oleda.Fill(ds, "Policies");

                    foreach (var row in ds.Tables[0].DefaultView)
                    {

                        var func = ((System.Data.DataRowView)row).Row.ItemArray[1].ToString();

                        if (func == function)
                        {
                            conRef = ((System.Data.DataRowView)row).Row.ItemArray[0].ToString();
                        }
                    }




                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }

            return conRef;

        }
        public Decimal getPremuimFromRateTable(string age, string rolePlayer, string sumAsured, string product)
        {
            _connString = "Provider= Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:/Users/G992107/Documents/GitHub/ILR_TestSuite/ILR_TestSuite/MIP UAT Test Scenarios/Rate_Table_Safrican Just Funeral.xlsx" + ";Extended Properties='Excel 8.0;HDR=Yes'";
            var premium="";
            
            using (OleDbConnection conn = new OleDbConnection(_connString))
            {
                try
                {

                    var cover = rolePlayer + "_" + sumAsured;
                    // Open connection
                    conn.Open();
                    string cmdQuery = $"SELECT ALB, {cover} FROM [{product}$]";

                    OleDbCommand cmd = new OleDbCommand(cmdQuery, conn);

                    // Create new OleDbDataAdapter
                    OleDbDataAdapter oleda = new OleDbDataAdapter();

                    oleda.SelectCommand = cmd;

                    // Create a DataSet which will hold the data extracted from the worksheet.
                    DataSet ds = new DataSet();

                    // Fill the DataSet from the data extracted from the worksheet.
                    oleda.Fill(ds, "Policies");



                    foreach (var row in ds.Tables[0].DefaultView)
                    {

                        var excellAge = ((System.Data.DataRowView)row).Row.ItemArray[0].ToString();
                        if (excellAge == age)
                        {
                            premium = ((System.Data.DataRowView)row).Row.ItemArray[1].ToString();
                        }


                    }

                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
                

            }
            return Convert.ToDecimal(premium);




        }


        public void writeResultsToExcell(string results, string sheet , string function)
        {

            string connString = "Provider= Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:/Users/G992107/Documents/GitHub/ILR_TestSuite/ILR_TestSuite/MIP UAT Test Scenarios/Rate_Table_Safrican Just Funeral.xlsx" + ";Extended Properties='Excel 8.0;HDR=Yes'";


            using (OleDbConnection conn = new OleDbConnection(connString))
            {
                try
                {

                    // Open connection
                    conn.Open();
                    OleDbCommand cmd = conn.CreateCommand();

                    var testDate = DateTime.Now.ToString();

                    //Test_Date
                    cmd.CommandText = $"UPDATE [{sheet}$] SET Test_Date = '{testDate}' WHERE Function = '{function}';";
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = $"UPDATE [{sheet}$] SET Test_Results  = '{results}' WHERE Function = '{function}';";
                    cmd.ExecuteNonQuery();

                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }




        }


        public void DisconnectBrowser()

        {

            if (_driver != null)

                _driver.Quit();

        }



        public string GetSeleniumFormatTag(string inputControlName)

        {

            var result = $"//*[@id=\"{inputControlName}\"]";

            return result;

        }



        public void Delay(int delaySeconds)

        {

            Thread.Sleep(delaySeconds * 1000);

        }



    }

}