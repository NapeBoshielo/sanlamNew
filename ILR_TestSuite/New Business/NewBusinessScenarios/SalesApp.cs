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
           Delay(30);
          //  Product1000MinMaxAge();
            Delay(2);
          Product2000MinMaxAge();
         //Delay(2);
           // Product3000MinMaxAge();
            //Delay(10);
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


            Delay(3);
            IWebElement worksite = _driver.FindElement(By.XPath("//*[@id='downshift-1-input']"));
            worksite.SendKeys("Nike");
            worksite.SendKeys(Keys.ArrowDown);
            worksite.SendKeys(Keys.Enter);
            Delay(2);
            IWebElement employer = _driver.FindElement(By.XPath("//*[@id='downshift-2-input']"));
            //*[@id="downshift-1-input"]
            employer.SendKeys("Nike");
            employer.SendKeys(Keys.ArrowDown);
            employer.SendKeys(Keys.Enter);
            Delay(2);
            IWebElement yes = _driver.FindElement(By.XPath("/html/body/reach-portal/div/div/div/div[4]/div/div[2]/div/label[1]"));
            yes.Click();

            IWebElement cont = _driver.FindElement(By.XPath(" /html/body/reach-portal/div/div/div/div[5]/button[2]"));
            cont.Click();

            Delay(2);
            IWebElement agree = _driver.FindElement(By.XPath(" /html/body/reach-portal/div/div/div/div[2]/button"));
            agree.Click();

            //Personal Details
            Delay(2);
            //firstname
            _driver.FindElement(By.XPath("//*[@id='/name']")).SendKeys("Thulani");
            Delay(2);
            //maiden name
            _driver.FindElement(By.XPath("//*[@id='/maiden-surname']")).SendKeys("Bongo");
            Delay(2);
            _driver.FindElement(By.XPath("//*[@id='/surname']")).SendKeys("Bongo");

            Delay(2);
            _driver.FindElement(By.XPath("//*[@id='/id-number']")).SendKeys("7604146054087");

            IWebElement select = _driver.FindElement(By.XPath(" //*[@id='/ethnicity']"));
            SelectElement oselect = new SelectElement(select);
            oselect.SelectByIndex(1);

            IWebElement selectstatus = _driver.FindElement(By.XPath("//*[@id='/marital-status']"));
            SelectElement cselect = new SelectElement(selectstatus);
            cselect.SelectByIndex(1);

            _driver.FindElement(By.XPath("//*[@id='/contact-number']")).SendKeys("0675344558");


            _driver.FindElement(By.XPath("//*[@id='/email']")).SendKeys("nape@gmilco.za");


            IWebElement electstatus = _driver.FindElement(By.XPath("//*[@id='/nationality']"));
            SelectElement elect = new SelectElement(electstatus);
            elect.SelectByIndex(1);


            IWebElement lectstatus = _driver.FindElement(By.XPath("//*[@id='/country-of-birth']"));
            SelectElement lect = new SelectElement(lectstatus);
            lect.SelectByIndex(1);


            _driver.FindElement(By.XPath("//*[@id='/gross-monthly-income']")).SendKeys("10000");


            IWebElement PermanentEmp = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/section/form/div/div[16]/div/label[1]"));
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

            Delay(3);

            for (int i = 1; i < 5; i++)
            {
                _driver.FindElement(By.XPath($"//*[@id='gatsby-focus-wrapper']/article/section/form/div[1]/section[{i.ToString()}]/label")).Click();

            }
            //click next
            Delay(3);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();
            Delay(3);

            for (int i = 1; i < 5; i++)
            {
                _driver.FindElement(By.XPath($"//*[@id='gatsby-focus-wrapper']/article/section/form/div[1]/section[{i.ToString()}]/label/section/div[1]")).Click();
                
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
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();


            //click tickbox
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section/div[2]/input[1]")).Click();



            //click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();

            //click tickbox product
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/section/div[1]/div[3]/button/span")).Click();

            //click tickbox product
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/div/div[2]/label/div")).Click();
            

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

            IWebElement sliderbar = _driver.FindElement(By.ClassName("slider"));
            int widthslider = sliderbar.Size.Width;
            Delay(1);
            IWebElement slider = _driver.FindElement(By.ClassName("slider"));
            Actions slideraction = new Actions(_driver);
            slideraction.ClickAndHold(slider);
            slideraction.MoveByOffset(75, 0).Build().Perform();


            //Spouse

            //click Add 
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/button")).Click();


            //select relationship
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[4]/div[2]/div[1]/div/label[2]")).Click();

            //FirstName
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[1].name")).SendKeys("Lindokuhle");
            //Surname
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[1].surname")).SendKeys("Matlou");

            //ID Number
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[1].id-number")).SendKeys("8811205596085");

            //Cellphone
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[1].contact-number")).SendKeys("0679774589");


            //Cover Amount 



            IWebElement sliderbar1 = _driver.FindElement(By.ClassName("slider"));
            int widthslider2 = sliderbar1.Size.Width;
            Delay(1);
            IWebElement slider1 = _driver.FindElement(By.ClassName("slider"));
            Actions slideraction1 = new Actions(_driver);
            slideraction.ClickAndHold(slider1);
            slideraction.MoveByOffset(0, 0).Build().Perform();


            //child

            //click Add 
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/button")).Click();


            //select relationship
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[5]/div[2]/div[1]/div/label[3]")).Click();

            

            //FirstName
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[2].name")).SendKeys("Zodwa");
            //Surname
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[2].surname")).SendKeys("Matlou");

            //ID Number
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[2].id-number")).SendKeys("0711202540086"); 

            //Cellphone
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[2].contact-number")).SendKeys("0679774522");


            //Cover Amount 



            IWebElement sliderbar2 = _driver.FindElement(By.ClassName("slider"));
            int widthslider1 = sliderbar2.Size.Width;
            Delay(1);
            IWebElement slider2 = _driver.FindElement(By.ClassName("slider"));
            Actions slideraction2 = new Actions(_driver);
            slideraction.ClickAndHold(slider2);
            slideraction.MoveByOffset(0, 0).Build().Perform();



            //parent

            //click Add 
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/button")).Click();


            //select relationship
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[6]/div[2]/div[1]/div/label[4]")).Click();

            //FirstName
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[3].name")).SendKeys("Modli");
            //Surname
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[3].surname")).SendKeys("Matlou");

            //ID Number
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[3].id-number")).SendKeys("6711202684086");

            //Cellphone
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[3].contact-number")).SendKeys("0677654589");


            //Cover Amount 



            IWebElement sliderbar3 = _driver.FindElement(By.ClassName("slider"));
            int widthslider3 = sliderbar3.Size.Width;
            Delay(1);
            IWebElement slider3 = _driver.FindElement(By.ClassName("slider"));
            Actions slideraction3 = new Actions(_driver);
            slideraction.ClickAndHold(slider3);
            slideraction.MoveByOffset(0, 0).Build().Perform();

            //Extended

            //click Add 
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/button")).Click();


            //select relationship
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[7]/div[2]/div[1]/div/label[5]")).Click();


            //Extended Relationship Type

            IWebElement RelationshipType = _driver.FindElement(By.XPath("//*[@id='downshift-7-input']"));
            RelationshipType.SendKeys("Cousin");
            RelationshipType.SendKeys(Keys.ArrowDown);
            RelationshipType.SendKeys(Keys.Enter);


            //FirstName
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[4].name")).SendKeys("Zondi");
            //Surname
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[4].surname")).SendKeys("Zwane");

            //ID Number
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[4].id-number")).SendKeys("9409060029083");

            //Cellphone
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[4].contact-number")).SendKeys("0670974589");


            //Cover Amount 



            IWebElement sliderbar4 = _driver.FindElement(By.ClassName("slider"));
            int widthslider4 = sliderbar4.Size.Width;
            Delay(1);
            IWebElement slider4 = _driver.FindElement(By.ClassName("slider"));
            Actions slideraction4 = new Actions(_driver);
            slideraction4.ClickAndHold(slider4);
            slideraction4.MoveByOffset(100, 0).Build().Perform();

            //Click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();
            

            //payment reciever(Beneficiary)


            //click relationship 
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/div/section/div[3]/div/div[1]/div/label[1]")).Click();



            //FirstName
            Delay(1);
            _driver.FindElement(By.Name("/funeral-beneficiaries[0].name")).SendKeys("Lindokuhle");
            //Surname
            Delay(1);
            _driver.FindElement(By.Name("/funeral-beneficiaries[0].surname")).SendKeys("Matlou");

            //ID Number
            Delay(1);
            _driver.FindElement(By.Name("/funeral-beneficiaries[0].id-number")).SendKeys("8811205596085");

            //Cellphone
            Delay(1);
            _driver.FindElement(By.Name("/funeral-beneficiaries[0].contact-number")).SendKeys("0679774589");

            // Percentage
            IWebElement sliderbar5 = _driver.FindElement(By.ClassName("slider"));
            int widthslider5 = sliderbar5.Size.Width;
            Delay(1);
            IWebElement slider5 = _driver.FindElement(By.ClassName("slider"));
            Actions slideraction5 = new Actions(_driver);
            slideraction5.ClickAndHold(slider5);
            slideraction5.MoveByOffset(260,0).Build().Perform();


            //Click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();




            //Word of advice
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='/record-of-advice']")).SendKeys("Test");
            

            //Click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();

            //Click No
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/section/div[2]/form/div[2]/div/label[2]")).Click();
            

            //go to payment 
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div/a[2]")).Click();





            /////////Payment Details
            string Bank = "";

            //policy payer
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='/same-as-fna']")).Click();


            //bank details
            //Bank Selction
            var value = "";
            switch (Bank)
            {
                case "ABSA BANK":
                    value = "ABSA BANK";
                    break;
                case "AFRICAN BANK":
                    value = "AFRICAN BANK";
                    break;

                case "BIDVEST BANK":
                    value = "BIDVEST BANK";
                    break;

                case "CAPITEC BANK":
                    value = "CAPITEC BANK";
                    break;
                case "DISCOVERY BANK LIMITED":
                    value = "DISCOVERY BANK LIMITED";
                    break;

                case "FIRST NATIONAL BANK":
                    value = "FIRST NATIONAL BANK";
                    break;

                case "STANDARD BANK S.A.":
                    value = "STANDARD BANK S.A.";
                    break;
                case "NEDBANK LIMITED":
                    value = "NEDBANK LIMITED";
                    break;

                case "TYME BANK":
                    value = "TYME BANK";
                    break;

                default:
                    break;
            }

            SelectElement oSelect = new SelectElement(_driver.FindElement(By.Name("/bank-select")));
            oSelect.SelectByValue(value);

            //Account Number
            Delay(1);
            _driver.FindElement(By.Name("")).SendKeys("");


            //Account Type
            Delay(1);
            _driver.FindElement(By.XPath("")).Click();


            //Debit day
            Delay(1);
            _driver.FindElement(By.Name("")).SendKeys("");


            //Select Preferred Debit Order DayField is required.


            //Next button




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
