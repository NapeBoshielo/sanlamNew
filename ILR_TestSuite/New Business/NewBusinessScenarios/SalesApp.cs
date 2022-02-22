using ILR_TestSuite;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OpenQA.Selenium.Firefox;
using System.Data.OleDb;
using System.Data;

namespace ILR_TestSuite.New_Business.Sales_App
{

    [TestFixture]
    public class SalesApp : TestBase_NB
    {
        private IWebDriver _driver;
        [SetUp]
            public void startBrowser()

            {

                _driver = base.SiteConnection();


            }



            [Test, Order(1)]
            public void RunTest()
        {
            Delay(2);
            Product1000MinMaxAge();
            Delay(2);
            Product2000MinMaxAge();
            Delay(2);
            Product3000MinMaxAge();
            Delay(10);
        }

        [Category(" Product1000MinMaxAge")]
        public void Product1000MinMaxAge()
        {
            
            try
            {

             

            }
            catch (Exception ex)
            {
                DisconnectBrowser();
                throw ex;
            }
        }


        [Category(" Product2000MinMaxAge")]
        public void Product2000MinMaxAge()
        {
            try
            {
            

            }
            catch (Exception ex)
            {
                DisconnectBrowser();
                throw ex;
            }
        }

        public void Product3000MinMaxAge()
        {
            try
            {
            ");


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
