using NUnit.Framework;

using OpenQA.Selenium;

using OpenQA.Selenium.Chrome;

using System.Threading;

using OpenQA.Selenium.Support.UI;

using System.IO;

using System;

using OpenQA.Selenium.Firefox;




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
         

            _driver.Url = "https://ilr-ppe.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrppe/run.w?";

            _userName = "G992127";//TODO add your user name and password

            _password = "P@$$word47";

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