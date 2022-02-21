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



        }

          
        
    }
}
