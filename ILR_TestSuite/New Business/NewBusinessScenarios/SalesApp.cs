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
using OpenQA.Selenium.Interactions;

namespace ILR_TestSuite.New_Business.Sales_App
{

    [TestFixture]
    public class SalesApp : TestBase_NB

    {
        private IWebDriver _driver;
        private string sheet;
        [SetUp]
            public void startBrowser()

            {

                _driver = base.SiteConnection();
            sheet = "";

            }



            [Test, Order(1)]
            public void RunTest()
        {
           Delay(20);
          //  Product1000MinMaxAge();
          //  Delay(2);
          //  Product2000MinMaxAge();
         //   Delay(2);
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
                using (OleDbConnection conn = new OleDbConnection(base._connString))
                {

                    conn.Open();

                    String cmd = "SELECT * FROM []";
                
                
                
                
                
                
                }


            }
            catch (Exception ex)
            {
               
                throw ex;
            }
            //*[@id="gatsby-focus-wrapper"]/article/div[2]/div[1]/button

            _driver.SwitchTo().ActiveElement();
                _driver.FindElement(By.XPath("//*[@id='___gatsby']")).Click();
                Delay(5);
                IWebElement new_client = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/div[2]/div[1]/button"));
                new_client.Click();
                //  Actions action = new Actions(_driver);
                // action.MoveToElement(new_client).Perform()
                Delay(2);
                IWebElement town = _driver.FindElement(By.XPath("//*[@id='downshift-0-input']"));
               town.SendKeys("Johan");
               town.SendKeys(Keys.ArrowDown);
                town.SendKeys(Keys.Enter);


                Delay(2);
                IWebElement worksite =_driver.FindElement(By.XPath("//*[@id='downshift-1-input']"));
                worksite.SendKeys("health");
                worksite.SendKeys(Keys.ArrowDown);
                worksite.SendKeys(Keys.Enter);
                Delay(2);
                IWebElement employer = _driver.FindElement(By.XPath("//*[@id='downshift-2-input']"));

               employer.SendKeys("sap");
                employer.SendKeys(Keys.ArrowDown);
                employer.SendKeys(Keys.Enter);
                Delay(2);
                IWebElement yes =  _driver.FindElement(By.XPath("/html/body/reach-portal/div/div/div/div[4]/div/div[2]/div/label[1]"));
                yes.Click();

               IWebElement cont = _driver.FindElement(By.XPath(" /html/body/reach-portal/div/div/div/div[5]/button[2]"));
                cont.Click();

                Delay(2);
                IWebElement agree = _driver.FindElement(By.XPath(" /html/body/reach-portal/div/div/div/div[2]/button"));
                agree.Click();

                //Personal Details
                Delay(2);
                //firstname
                _driver.FindElement(By.XPath("//*[@id='/name']")).SendKeys("Nape");
                Delay(2);
                //maiden name
                _driver.FindElement(By.XPath("//*[@id='/maiden-surname']")).SendKeys("Bongo");
                Delay(2);
                _driver.FindElement(By.XPath("//*[@id='/surname']")).SendKeys("Bongo");

                Delay(2);
                _driver.FindElement(By.XPath("//*[@id='/id-number']")).SendKeys("0005297455080");

                IWebElement select = _driver.FindElement(By.XPath(" //*[@id='/ethnicity']"));
                SelectElement oselect = new SelectElement(select);
                oselect.SelectByIndex(1);

                IWebElement selectstatus = _driver.FindElement(By.XPath("//*[@id='/marital-status']"));
                SelectElement cselect = new SelectElement(selectstatus);
                cselect.SelectByIndex(1);

                _driver.FindElement(By.XPath("//*[@id='/contact-number']")).SendKeys("0823344558");


                _driver.FindElement(By.XPath("//*[@id='/email']")).SendKeys("nape@gmilco.za");


                IWebElement electstatus = _driver.FindElement(By.XPath("//*[@id='/nationality']"));
                SelectElement elect = new SelectElement(electstatus);
                elect.SelectByIndex(1);


                IWebElement lectstatus = _driver.FindElement(By.XPath("//*[@id='/country-of-birth']"));
                SelectElement lect = new SelectElement(lectstatus);
                lect.SelectByIndex(1);


                _driver.FindElement(By.XPath("//*[@id='/gross-monthly-income']")).SendKeys("10000");


         IWebElement PermanentEmp =  _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/section/form/div/div[16]/div/label[1]"));
            PermanentEmp.Click();


            IWebElement salaryF = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/section/form/div/div[17]/div/label[2]"));
       salaryF.Click();
            Delay(1);
            _driver.FindElement(By.XPath(" //*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a")).Click();


            //occupation
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/form/section/div/div[1]/label")).Click();
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[3]/div[1]/a[2]")).Click();

            ///dependants

            Delay(4);

            for (int i =1 ; i < 5; i++)
            {
                _driver.FindElement(By.XPath($"//*[@id='gatsby-focus-wrapper']/article/section/form/div[1]/section[{i.ToString()}]/label")).Click();

            }

            Delay(1);
            _driver.FindElement(By.XPath(" //*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();
            //save and continue
           // Delay(1);
          //  _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/nav/div/a")).Click();


        }


        [TearDown]
        public void closeBrowser()
        {
            base.DisconnectBrowser();
        }

    }
}
