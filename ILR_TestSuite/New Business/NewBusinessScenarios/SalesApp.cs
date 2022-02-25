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
   
        [SetUp]
            public void startBrowser()

            {

                _driver = base.SiteConnection();
           

            }



            [Test, Order(1)]
            public void RunTest()
        {
           Delay(28);
          //  Product1000MinMaxAge();
            Delay(2);
          Product2000MinMaxAge();
         //Delay(2);
           // Product3000MinMaxAge();
            //Delay(10);
        }
        public void createNewClient()
        {
            var policyHolderData = getPolicyHolderDetails("1");
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
            Delay(4);
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
            if (policyHolderData["permanently_employed"] == "Yes")
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
            Delay(1);
            _driver.FindElement(By.XPath(" //*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a")).Click();


            //occupation
            Delay(3);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/form/section/div/div[1]/label")).Click();
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[3]/div[1]/a[2]")).Click();

            ///dependants

            Delay(4);

            for (int i = 1; i < 5; i++)
            {
                _driver.FindElement(By.XPath($"//*[@id='gatsby-focus-wrapper']/article/section/form/div[1]/section[{i.ToString()}]/label")).Click();

            }
            //click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();

            //sclick on non applicable 
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/div[1]/div/button")).Click();


            //Net Salary After Deductions
            Delay(1);
            _driver.FindElement(By.Name("/total-salary-after-deductions")).SendKeys("50000");

            //Additional income

            Delay(1);
            _driver.FindElement(By.Name("/additional-income")).SendKeys("5000");


            //Existing Financial Cover
            Delay(1);
            _driver.FindElement(By.Name("/existing-financial-cover")).SendKeys("0");

            //School Fees
            Delay(1);
            _driver.FindElement(By.Name("/school-fees")).SendKeys("4000");

            //Food
            Delay(1);
            _driver.FindElement(By.Name("/food")).SendKeys("3000");
            //Retail accounts
            Delay(1);
            _driver.FindElement(By.Name("/retail-accounts")).SendKeys("1000");

            //Cellphone
            Delay(1);
            _driver.FindElement(By.Name("/cellphone")).SendKeys("700");
            //Debt
            Delay(1);
            _driver.FindElement(By.Name("/debt")).SendKeys("0");

            // Mortgage / Rent
            Delay(1);
            _driver.FindElement(By.Name("/mortage-rent")).SendKeys("9000");
            //Transport
            Delay(1);
            _driver.FindElement(By.Name("/transport")).SendKeys("2500");
            //Entertainment / Other
            Delay(1);
            _driver.FindElement(By.Name("/entertainment-other")).SendKeys("3000");



            //click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/button")).Click();

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


        public void Product2000MinMaxAge()
        {

            createNewClient();
            //click tickbox
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section/div[2]/input[1]")).Click();



            //click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();

            //click tickbox product
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/section/div[1]/div[2]/button/span")).Click();

            //click tickbox product
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/div/div[2]/label/divn")).Click();


            //click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a")).Click();



            //click on No
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[1]/div/div[2]/div/div/label[2]")).Click();


            //click on 5%
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[2]/div/div[2]/div/div/label[1]")).Click();


            //add main life
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[3]/div[2]/div[1]/div/label[1]")).Click();

            //Cover Amount 

        
            IWebElement sliderbar = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[4]/div[4]/div[1]"));
            int widthslider = sliderbar.Size.Width;
            Delay(1);

            IWebElement slider = _driver.FindElement(By.ClassName("slinder"));

            Actions slideraction = new Actions(_driver);
            slideraction.ClickAndHold(slider);
            slideraction.MoveByOffset(40, 0).Build().Perform();


            //Spouse

            //FirstName
            Delay(1);
            _driver.FindElement(By.XPath("")).SendKeys("");
            //Surname
            Delay(1);
            _driver.FindElement(By.XPath("")).SendKeys("");

            //ID Number
            Delay(1);
            _driver.FindElement(By.XPath("")).SendKeys("");

            //Cellphone
            Delay(1);
            _driver.FindElement(By.XPath("")).SendKeys("");




            //Cover Amount 





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
