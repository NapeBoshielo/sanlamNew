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
using DocumentFormat.OpenXml.Spreadsheet;

namespace ILR_TestSuite.New_Business.Sales_App
{

    [TestFixture]
    public class SalesApp : TestBase_NB

    {
        private string results;

        /// IWebElement main_="",spouce_,child_="",parent_="",extended_="";
        //  string main1_="";
        [SetUp]
            public void startBrowser()

            {

                _driver = base.SiteConnection();
            string sheet = "Scenarios";



        }



            [Test, Order(1)]
            public void RunTest()
        {
            Delay(35);

            createNewClient();
            using (OleDbConnection conn = new OleDbConnection(_test_data_connString))
            {
                try
                {

                    // Open connection
                    conn.Open();
                    string cmdQuery = "SELECT * FROM [Scenarios$]";

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
                        var scenarioID = ((System.Data.DataRowView)row).Row.ItemArray[0].ToString(); ;
                        var func = ((System.Data.DataRowView)row).Row.ItemArray[4].ToString();
                        var product = ((System.Data.DataRowView)row).Row.ItemArray[1].ToString();

                        switch (func)
                        {
                            case "ChildMaxAge":
                                ChildMaxAge();
                                break;
                            case "ExtendedMaxAge":
                                ExtendedMaxAge();
                                break;
                            default:
                                break;
                        }
                        //validation(scenarioID,func,product)
                    }
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }





        }

        private void maxSpouse()
        {
            throw new NotImplementedException();
        }

        public void createNewClient()
        {   //get policy holder data
            var policyHolderData = getPolicyHolderDetails("1");
            _driver.SwitchTo().ActiveElement();
            _driver.FindElement(By.XPath("//*[@id='___gatsby']"));
            Delay(2);
            IWebElement new_client = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/div[2]/div[1]/button"));
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
            IWebElement yes = _driver.FindElement(By.XPath("/html/body/reach-portal/div/div/div/div[4]/div/div[2]/div/label[1]"));
            Delay(1);
            yes.Click();
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

            Delay(2);

            _driver.FindElement(By.XPath("//*[@id='/id-number']")).SendKeys("1803148913086");
            

            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/button")).Click();
            Delay(1);
            //Min validation
            var Errormessage = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/section/form/div/div[8]/span/span")).Text;
            {
                TakeScreenshot(_driver, $@"{_screenShotFolder}\Failed_Scenarios\", "MinMainLifeValidation");
            }

            if ((Errormessage) == ("Must be at least 18 years old."))
            {
                results = "Passed";
            }
            else
            {
                results = "Failed";
            }

            // base.writeResultsToExcell(results, sheet, "MainLifeMinAge");

            Delay(2);



            //click field
            Actions builder = new Actions(_driver);
            builder.MoveToElement(_driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/section/form/div/div[4]/input")))
            .Click().Build().Perform();
           
            //claer  keys
            IWebElement clear = _driver.FindElement(By.XPath("//*[@id='/id-number']"));
            clear.SendKeys(Keys.Control + "a");
            clear.SendKeys(Keys.Delete);
            Delay(3);

           //Send new keys
            IWebElement element = _driver.FindElement(By.XPath("//*[@id='/id-number']"));
            Delay(3);
            element.SendKeys("3403148170080");
            Delay(2);

            //Min validation
            var Errormessage2 = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/section/form/div/div[8]/span/span")).Text;
            {
                TakeScreenshot(_driver, $@"{_screenShotFolder}\Failed_Scenarios\", "MaxMainLifeValidation");
            }

            if ((Errormessage2) == ("Must not be older than 74 years of age."))
            {
                results = "Passed";
            }
            else
            {
                results = "Failed";
            }



            // base.writeResultsToExcell(results, sheet, "MainLifeMaxAge");





            //click field
            Actions builder2 = new Actions(_driver);
            builder2.MoveToElement(_driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/section/form/div/div[4]/input")))
            .Click().Build().Perform();

            //clear keys
            IWebElement clear2 = _driver.FindElement(By.XPath("//*[@id='/id-number']"));
            clear2.SendKeys(Keys.Control + "a");
            clear2.SendKeys(Keys.Delete);
            Delay(3);
            //send new keys
            IWebElement element1 = _driver.FindElement(By.XPath("//*[@id='/id-number']"));
            Delay(3);
            element1.SendKeys(policyHolderData["idNo"]);
            Delay(2);

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
            _driver.FindElement(By.Name("/total-salary-after-deductions")).SendKeys(policyHolderData["net_salary"]);

            //Additional income

            Delay(1);
            _driver.FindElement(By.Name("/additional-income")).SendKeys(policyHolderData["additional_income"]);


            //Existing Financial Cover
            Delay(1);
            _driver.FindElement(By.Name("/existing-financial-cover")).SendKeys(policyHolderData["existing_financial_cover"]);

            //School Fees
            Delay(1);
            _driver.FindElement(By.Name("/school-fees")).SendKeys(policyHolderData["school_fees"]);

            //Food
            Delay(1);
            _driver.FindElement(By.Name("/food")).SendKeys(policyHolderData["food"]);
            //Retail accounts
            Delay(1);
            _driver.FindElement(By.Name("/retail-accounts")).SendKeys(policyHolderData["retail_accounts"]);

            //Cellphone
            Delay(1);
            _driver.FindElement(By.Name("/cellphone")).SendKeys(policyHolderData["cellphone"]);
            //Debt
            Delay(1);
            _driver.FindElement(By.Name("/debt")).SendKeys(policyHolderData["debt"]);

            // Mortgage / Rent
            Delay(1);
            _driver.FindElement(By.Name("/mortage-rent")).SendKeys(policyHolderData["mortage"]);
            //Transport
            Delay(1);
            _driver.FindElement(By.Name("/transport")).SendKeys(policyHolderData["transport"]);
            //Entertainment / Other
            Delay(1);
            _driver.FindElement(By.Name("/entertainment-other")).SendKeys(policyHolderData["entertainment_other"]);


            //click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();


            //click tickbox for agreement 
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section/div[2]/input[1]")).Click();



            //click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();

 


        }
        [Category("ChildMaxAge")]
        public void ChildMaxAge()
        {


            string results = "";


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

            //select relationship
            Delay(1);

            // Locate the element “child” by By.xpath. 
            IWebElement Child = _driver.FindElement(By.Name("/cover-details[0].relationship-type"));

            // Create an object of Actions class and pass reference variable driver as a parameter to its constructor. 
            Actions actions = new Actions(_driver);
            actions.Click(Child);
            actions.Build().Perform();


            //FirstName
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[0].name")).SendKeys("Zodwa");
            //Surname
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[0].surname")).SendKeys("Matlou");

            //ID Number
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[0].id-number")).SendKeys("7001113329081");

            //Cellphone
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[0].contact-number")).SendKeys("0679774522");


            //select relationship
            Delay(2);
            //Cover Amount 
            SlideBar("Child");

            TakeScreenshot(_driver, $@"{_screenShotFolder}\Failed_Scenarios\", "MaxChildValidation");

            //click next 
            Delay(2);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/button")).Click();

            Delay(2);


            //Max validation
            var Errormessage = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[3]/div[4]/div[2]/span")).Text;


            if ((Errormessage) != ("No cover is available for the entered details"))
            {
                results = "Passed";
            }
            else
            {
                results = "Failed";
            }



            // base.writeResultsToExcell(results, sheet, "ChildMaxAge");

        }

        [Category("ExtendedMaxAge")]
        public void ExtendedMaxAge()
        {


            string results = "";
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

            //select relationship
            Delay(1);




            //Extended

            //select relationship
            // Locate the element “extended” by By.xpath. 
            IWebElement extended = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[3]/div[2]/div[1]/div/label[5]"));

            // Create an object of Actions class and pass reference variable driver as a parameter to its constructor. 
            Actions actions = new Actions(_driver);
            actions.Click(extended);
            actions.Build().Perform();


            //Extended Relationship Type

            IWebElement RelationshipType = _driver.FindElement(By.Id("downshift-3-input"));
            RelationshipType.SendKeys("Cousin");
            RelationshipType.SendKeys(Keys.ArrowDown);
            RelationshipType.SendKeys(Keys.Enter);


            //FirstName
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[0].name")).SendKeys("Zondi");
            //Surname
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[0].surname")).SendKeys("Zwane");

            //ID Number
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[0].id-number")).SendKeys("3011148884087");

            //Cellphone
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[0].contact-number")).SendKeys("0670974589");


            //Cover Amount 

            SlideBar("Extended");

           

            TakeScreenshot(_driver, $@"{_screenShotFolder}\Failed_Scenarios\", "ExtendedMaxvalidation");

            //click next 
            Delay(2);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/button")).Click();

            Delay(2);


            //Max validation
            var Errormessage = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[3]/div[4]/div[1]/label")).Text;


            if ((Errormessage) == ("Cover is only available for persons up to 85 years of age"))
            {
                results = "Passed";
            }
            else
            {
                results = "Failed";
            }


            // base.writeResultsToExcell(results, sheet, "ExtendedMaxAge");

        }


        public void SlideBar(string roles)
        {

          string Amount = "";

            using (OleDbConnection con = new OleDbConnection(base._test_data_connString))
            {
                try
                {
                    con.Open();

                    String command = "SELECT * FROM [CoverAmount$]";

                    OleDbCommand cmd = new OleDbCommand(command, con);

                    OleDbDataAdapter adapt = new OleDbDataAdapter();
                    adapt.SelectCommand = cmd;

                    DataSet ds = new DataSet("policies");
                    adapt.Fill(ds);
                    foreach (var row in ds.Tables[0].DefaultView)
                    {

                    var  role = ((System.Data.DataRowView)row).Row.ItemArray[0].ToString();
                        
                        if (role==roles) {
                        
                        roles = ((System.Data.DataRowView)row).Row.ItemArray[0].ToString();
                            Amount = ((System.Data.DataRowView)row).Row.ItemArray[1].ToString();
                            break;
                        }

                       
                    }

                }
                catch (Exception ex)

                {
                    throw ex;

                }
                con.Close();
                con.Dispose();

            }
            if (roles == "Myself" || roles == "Child" || roles == "Spouce" )
            {

                var V_Position = "";
                switch (Amount)
                {

                    case "5000":
                        V_Position = "0";
                        break;
                    case "7000":
                        V_Position = "12.5";
                        break;
                    case "10000":
                        V_Position = "25";
                        break;
                    case "15000":
                        V_Position = "37.5";
                        break;
                    case "20000":
                        V_Position = "50";
                        break;

                    case "30000":
                        V_Position = "62.5";
                        break;
                    case "40000":
                        V_Position = "75";
                        break;
                    case "50000":
                        V_Position = "87.5";
                        break;
                    case "60000":
                        V_Position = "100";
                        break;

                }

                IWebElement sliderbar = _driver.FindElement(By.ClassName("slider"));
                int widthslider = sliderbar.Size.Width;
                Delay(1);
                IWebElement slider = _driver.FindElement(By.ClassName("slider"));
                Actions slideraction = new Actions(_driver);
                slideraction.ClickAndHold(slider);
                slideraction.MoveByOffset(Convert.ToInt32(V_Position), 0).Build().Perform();


            }
            else if (roles == "Parent")
            {

                var V_Position = "";
                switch (Amount)
                {

                    case "5000":
                        V_Position = "0";
                        break;
                    case "7500":
                        V_Position = "25";
                        break;
                    case "10000":
                        V_Position = "50";
                        break;
                    case "15000":
                        V_Position = "75";
                        break;
                    case "20000":
                        V_Position = "100";
                        break;

                }

                IWebElement sliderbar = _driver.FindElement(By.ClassName("slider"));
                int widthslider = sliderbar.Size.Width;
                Delay(1);
                IWebElement slider = _driver.FindElement(By.ClassName("slider"));
                Actions slideraction = new Actions(_driver);
                slideraction.ClickAndHold(slider);
                slideraction.MoveByOffset(Convert.ToInt32(V_Position), 0).Build().Perform();



            }
            else {
                var V_Position = "";
                switch (Amount)
                {

                    case "5000":
                        V_Position = "0";
                        break;
                    case "7500":
                        V_Position = "20";
                        break;
                    case "10000":
                        V_Position = "40";
                        break;
                    case "15000":
                        V_Position = "60";
                        break;
                    case "20000":
                        V_Position = "80";
                        break;
                    case "30000":
                        V_Position = "80";
                        break;
                }
                IWebElement sliderbar = _driver.FindElement(By.ClassName("slider"));
                int widthslider = sliderbar.Size.Width;
                Delay(1);
                IWebElement slider = _driver.FindElement(By.ClassName("slider"));
                Actions slideraction = new Actions(_driver);
                slideraction.ClickAndHold(slider);
                slideraction.MoveByOffset(Convert.ToInt32(V_Position), 0).Build().Perform();



            }




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
