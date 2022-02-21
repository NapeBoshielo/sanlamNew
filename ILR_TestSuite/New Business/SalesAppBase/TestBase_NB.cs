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

    public class TestBase_NB

    {

        private ChromeOptions _chromeOptions;

        public IWebDriver _driver;

        private string _userName;

        private string _password;

        private string _pin;

        public string _screenShotFolder;

        public string _connString;




        [SetUp]

        public void StartBrowser()
        {


            _chromeOptions = new ChromeOptions();
            _chromeOptions.AddArguments("--incognito");
            _chromeOptions.AddArguments("--ignore-certificate-errors");
            _driver = new ChromeDriver("C:/Code/bin", _chromeOptions);



            _connString = "Provider= Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:/Users/G992107/Documents/GitHub/ILR_TestSuite/ILR_TestSuite/MIP UAT Test Scenarios/TestData.xlsx" + ";Extended Properties='Excel 8.0;HDR=Yes'";

            _screenShotFolder = $@"C:\Users\G992107\Documents\GitHub\ILR_TestSuite\Failed_ScreenShots​{ScreenShotDailyFolderName()}​\";

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

            _driver.Url = "https://uat-fe.safricansalesapp.net/advisor/dashboard/";

            try
            {

               

                _userName = "G992127";//TODO add your user name and password

                _password = "P@$$word47";

                _pin = "119547";

                _driver.Manage().Window.Maximize();

                System.Threading.Thread.Sleep(2000);

                _driver.Manage().Window.Maximize();

                System.Threading.Thread.Sleep(2000);

                _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/div/div[2]/button")).Click();
                System.Threading.Thread.Sleep(3000);

                IWebElement loginTextBox = _driver.FindElement(By.Name("username"));
                System.Threading.Thread.Sleep(3000);
                IWebElement passwordTextBox = _driver.FindElement(By.Name("password"));
                System.Threading.Thread.Sleep(3000);
                IWebElement loginBtn = _driver.FindElement(By.XPath("/html/body/div[1]/form/div/button[1]"));
                System.Threading.Thread.Sleep(3000);
                loginTextBox.SendKeys(_userName);
                System.Threading.Thread.Sleep(3000);
                passwordTextBox.SendKeys(_password);
                System.Threading.Thread.Sleep(3000);
                loginBtn.Click();
                System.Threading.Thread.Sleep(3000);

                IWebElement pinTextBox = _driver.FindElement(By.XPath("/html/body/div/div[1]/div/section/form/div/input"));
                pinTextBox.SendKeys(_pin);
                System.Threading.Thread.Sleep(3000);

                _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div/section/form/button")).Click();
                


                return _driver;
            }
            catch (Exception ex) {

       


                throw ex;
            }

           
           
        }
        public string GetPolicyNoFromExcell(string sheet, string function)
        {


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
            var premium = "";

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


        public void writeResultsToExcell(string results, string sheet, string function)
        {
            using (OleDbConnection conn = new OleDbConnection(_connString))
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