using NUnit.Framework;

using static ILR_TestSuite.TestBase;

using OpenQA.Selenium;

using OpenQA.Selenium.Interactions;

using System;

using System.Collections.Generic;
using ILR_TestSuite;
using System.Linq;
using OpenQA.Selenium.Support.UI;

namespace PolicyServicing

{

    [TestFixture]

    public class PolicyServicing : TestBase
    {



        private IWebDriver _driver;

        [SetUp]

        public void startBrowser()

        {



            _driver = base.SiteConnection();

        }



        [Test, Order(1)]

        public void PolicyServicingTestSuite()

        {

            Delay(2);


            //LogInValidation();



            Delay(20);

        }

        [Category("LogInValidation")]
        public void policySearch( string contractRef = "")
        {
            var policyStatusCode = "";
            IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;

            Delay(2);
            //Click on Miain
            _driver.FindElement(By.Name("CBWeb")).Click();
            Delay(2);
            //Click on contract search 
            _driver.FindElement(By.Name("alf-ICF8_00000222")).Click();
            Delay(2);


            //Click on product
            _driver.FindElement(By.Name("frmProductCode")).Click();
            Delay(2);

            //Type in contract ref if there is any 
      
              _driver.FindElement(By.Name("frmContractReference")).SendKeys(contractRef);


          
            Delay(4);



        





            //Click on Search Icon 
            _driver.FindElement(By.Name("btncbcts0")).Click();
            Delay(2);

        }

     

        [Category("LogInValidation")]

        public void LogInValidation()

        {

            try

            {
                //Arrange
                var expectedWelcomePage = "Welcome to Sanlam ARL Demo for the Web";
               

                //Action

                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/table/tbody/tr[2]/td[3]/center[1]/b"));

               TakeScreenshot(_driver, $@"{_screenShotFolder}\LogInValidation\", "Landing Page Validation");

                //find the result
                string actualResult = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/table/tbody/tr[2]/td[3]/center[1]/b")).Text;

                Assert.IsTrue(expectedWelcomePage.Equals(actualResult, StringComparison.CurrentCultureIgnoreCase));


            }

            catch (Exception ex)

            {

                DisconnectBrowser();

                throw ex;

            }

        }
        [TearDown]

        public void closeBrowser()

        {

            base.DisconnectBrowser();

        }



    }

}
