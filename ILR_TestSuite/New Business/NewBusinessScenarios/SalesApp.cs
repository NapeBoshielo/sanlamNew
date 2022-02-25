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
           Delay(40);
           Product1000MinMaxAge();
          //  Delay(2);
          //  Product2000MinMaxAge();
         //   Delay(2);
           // Product3000MinMaxAge();
            Delay(10);
        }

        [Category(" Product1000MinMaxAge")]
        public void Product1000MinMaxAge()
        {
            
            try
            {
                createNewClient();
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
        public void createNewClient()
        {
             var policyHolderData= getPolicyHolderDetails("1");
             _driver.SwitchTo().ActiveElement();
             _driver.FindElement(By.XPath("//*[@id='___gatsby']"));
             Delay(2);
             IWebElement new_client = _driver.FindElement(By.XPath(" //*[@id='gatsby-focus-wrapper']/article/div[2]/div[1]/button"));
             new_client.Click();
             //  Actions action = new Actions(_driver);
            // action.MoveToElement(new_client).Perform()
            Delay(2);
            IWebElement town = _driver.FindElement(By.XPath("//*[@id='downshift-0-input']"));
            town.SendKeys(policyHolderData["town"]);
            Delay(1);
            town.SendKeys(Keys.ArrowDown);
            Delay(1);
            town.SendKeys(Keys.Enter);
            Delay(2);
            IWebElement worksite = _driver.FindElement(By.XPath("//*[@id='downshift-1-input']"));
            worksite.SendKeys(policyHolderData["WorkSite"]);
            Delay(1);
            worksite.SendKeys(Keys.ArrowDown);
            Delay(1);
            worksite.SendKeys(Keys.Enter);
            Delay(2);
            IWebElement employer = _driver.FindElement(By.XPath("//*[@id='downshift-2-input']"));
            employer.SendKeys(policyHolderData["employemnt"]);
            Delay(1);
            employer.SendKeys(Keys.ArrowDown);
            Delay(1);
            employer.SendKeys(Keys.Enter);
            Delay(2);
            IWebElement no = _driver.FindElement(By.XPath("/html/body/reach-portal/div/div/div/div[4]/div/div[2]/div/label[2]"));
            Delay(1);
            no.Click();
            Delay(2);
            IWebElement cont = _driver.FindElement(By.XPath(" /html/body/reach-portal/div/div/div/div[5]/button[2]"));
            cont.Click();
            Delay(2);
            IWebElement agree = _driver.FindElement(By.XPath(" /html/body/reach-portal/div/div/div/div[2]/button"));
            agree.Click();
            //Personal Details
            Delay(2);
            //firstname
            _driver.FindElement(By.XPath("//*[@id='/name']")).SendKeys(policyHolderData["first_name"]);
            Delay(2);
            //maiden name
            _driver.FindElement(By.XPath("//*[@id='/maiden-surname']")).SendKeys(policyHolderData["maiden"]);
            Delay(2);
            _driver.FindElement(By.XPath("//*[@id='/surname']")).SendKeys(policyHolderData["surname"]);
            Delay(2);
           _driver.FindElement(By.XPath("//*[@id='/id-number']")).SendKeys(policyHolderData["idNo"]);
            Delay(2);
            //Select ethicity
           IWebElement select = _driver.FindElement(By.XPath(" //*[@id='/ethnicity']"));
           SelectElement oselect = new SelectElement(select);
           oselect.SelectByValue(policyHolderData["ethnicity"]);
           Delay(2);
           //Select Maratiel
           IWebElement selectstatus = _driver.FindElement(By.XPath("//*[@id='/marital-status']"));
           SelectElement cselect = new SelectElement(selectstatus);
           cselect.SelectByValue(policyHolderData["material_status"]);
           Delay(2);
           //Enter contact number
           _driver.FindElement(By.XPath("//*[@id='/contact-number']")).SendKeys(policyHolderData["cellphone_number"]);
           Delay(2);
           //Enter email
           _driver.FindElement(By.XPath("//*[@id='/email']")).SendKeys(policyHolderData["email"]);
            Delay(2);
           // //select nationality
           // IWebElement electstatus = _driver.FindElement(By.XPath("//*[@id='/nationality']"));
           // SelectElement elect = new SelectElement(electstatus);
           // elect.SelectByText(policyHolderData["nationality"]);
           // Delay(2);
           // //Country of birth
           //IWebElement lectstatus = _driver.FindElement(By.XPath("//*[@id='/country-of-birth']"));
           //SelectElement lect = new SelectElement(lectstatus);
           //lect.SelectByValue(policyHolderData["countryOfResidence"]);
           // Delay(2);
            //Enter gross monthly
           _driver.FindElement(By.XPath("//*[@id='/gross-monthly-income']")).SendKeys(policyHolderData["grossMonthlyIncome"]);
            Delay(2);

            //Select employent type
            if(policyHolderData["permanently_employed"] == "Yes")
            {
                _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/section/form/div/div[16]/div/label[1]")).Click();

            }
            else
            {
                _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/section/form/div/div[16]/div/label[2]")).Click();

            }
            Delay(2);

            //Salary frequency
            switch (policyHolderData["salary_frequency"])
            {
                case "Weekly":
                   _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/section/form/div/div[17]/div/label[1]")).Click();

                    break;
                case "Monthly":
                   _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/section/form/div/div[17]/div/label[2]")).Click();
                    break;

                case "Other":
                   _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/section/form/div/div[17]/div/label[3]")).Click();
                    break;

                default:
                   _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/section/form/div/div[17]/div/label[2]")).Click();
                    break;

                  
            }
            Delay(2);
            _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/div[2]/div[1]/button")).Click();
            Delay(2);

        }
        public void Product3000MinMaxAge()
        {
          
           
        }


        [TearDown]
        public void closeBrowser()
        {
            base.DisconnectBrowser();
        }

    }
}
