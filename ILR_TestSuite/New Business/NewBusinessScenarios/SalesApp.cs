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
       /// IWebElement main_="",spouce_,child_="",parent_="",extended_="";
      //  string main1_="";
        [SetUp]
            public void startBrowser()

            {

                _driver = base.SiteConnection();
           

            }



            [Test, Order(1)]
            public void RunTest()
        {
            Delay(35);
            var mlMaxMinAgeRes = createNewClient();
            writeResultsToExcell(mlMaxMinAgeRes, "Scenarios", "MaxMinAgeMainLife");

            using (OleDbConnection conn = new OleDbConnection(_test_data_connString))
            {
                try
                {
                    var sheet = "Scenarios";
                    var results ="";
                    // Open connection
                    conn.Open();
                    string cmdQuery = "SELECT * FROM ["+ sheet + "$]";

                    OleDbCommand cmd = new OleDbCommand(cmdQuery, conn);

                    // Create new OleDbDataAdapter
                    OleDbDataAdapter oleda = new OleDbDataAdapter();

                    oleda.SelectCommand = cmd;

                    // Create a DataSet which will hold the data extracted from the worksheet.
                    DataSet ds = new DataSet();

                    // Fill the DataSet from the data extracted from the worksheet.
                    oleda.Fill(ds, "Policies");


                    addMainLife();
                    foreach (var row in ds.Tables[0].DefaultView)
                    {
                        var func = ((System.Data.DataRowView)row).Row.ItemArray[3].ToString();
           
                            switch (func)
                            {
                                case "MaxMinAgeSpouse":
                                    results = SpouseMaxMin();
                                    writeResultsToExcell(results, sheet, func);
                                    break;
                                case "MaxMinAgeChild":
                                    results = ChildMaxAge();
                                    writeResultsToExcell(results, sheet, func);
                                    break;
                                case "MaxMinAgeParent":
                                    results = MaxMinAgeParent();
                                writeResultsToExcell(results, sheet, func);
                                    break;
                                case "MaxMinAgeExtended":
                                    results = ExtendedMaxAge();
                                    writeResultsToExcell(results, sheet, func);
                                    break;
                              

                            default:
                                    break;
                            }
                      
                     
                    }
                    

                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }





        }
        [Category("MaxMin Age for Parent")]
        public string MaxMinAgeParent()
        {



            string results = "";
            try
            {
                //click Add
                Delay(1);
                _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/button")).Click();


                //select relationship
                Delay(1);
                _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[4]/div[2]/div[1]/div/label[4]")).Click();


                //FirstName
                Delay(1);
                _driver.FindElement(By.XPath("//*[@id='/cover-details[3].name']")).SendKeys("Modli");
                //Surname
                Delay(1);
                _driver.FindElement(By.Name("/cover-details[3].surname")).SendKeys("Matlou");



                //ID Number
                Delay(1);
                _driver.FindElement(By.Name("/cover-details[3].id-number")).SendKeys("0003099707089");



                //Cellphone
                Delay(1);
                _driver.FindElement(By.Name("/cover-details[3].contact-number")).SendKeys("0677654589");




                String MinimumValidation = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[6]/div[4]/div[1]/label")).Text;

                if (MinimumValidation == "Cover is only available for parents from 26 to 85 years of age")
                {



                    TakeScreenshot(_driver, $@"{_screenShotFolder}\Failed_Scenarios\", "ParentMinValidation");
                    results = "Passed";



                }
                else
                {
                    results = "Failed";
                }
                Delay(1);



                //ID Number
                Delay(4);

                //click field
                Actions builder = new Actions(_driver);
                builder.MoveToElement(_driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/section/form/div/div[4]/input")))
                .Click().Build().Perform();

                //clear keys
                IWebElement clear = _driver.FindElement(By.XPath("/cover-details[1].id-number"));
                clear.SendKeys(Keys.Control + "a");
                clear.SendKeys(Keys.Delete);
                Delay(3);

                //Send new keys
                IWebElement element = _driver.FindElement(By.XPath("/cover-details[1].id-number"));
                Delay(3);
                element.SendKeys("3403148170080");
                Delay(2);





                //ID Number
                Delay(1);
                _driver.FindElement(By.Name("/cover-details[1].id-number")).SendKeys("3003095612082");





                // TakeScreenshot(_driver, $@"{_screenShotFolder}\Failed_Scenarios\", "ParentMax");
                String MaximumValidation = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[6]/div[4]/div[1]/label")).Text;



                if (MinimumValidation.Contains("Cover is only available for parents from 26 to 85 years of age") && MaximumValidation.Contains("Cover is only available for parents from 26 to 85 years of age"))

                {



                    results = "Passed";
                    TakeScreenshot(_driver, $@"{_screenShotFolder}\validations\", "ParentMaxValidation");



                }



                else
                {
                    results = "Failed";
                    TakeScreenshot(_driver, $@"{_screenShotFolder}\Failed_Scenarios\", "ParentMaxValidation");

                }




            }
            catch (Exception ex) { }
            return results;




        }
        [Category("Add Main Life")]
        public void addMainLife()
        {

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

            SlideBar("Myself");
        }

        [Category("ChildMaxAge")]
        public string ChildMaxAge()
        {


            string results = "";

            //Add life 
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/button")).Click();

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
            _driver.FindElement(By.Name("/cover-details[2].name")).SendKeys("Zodwa");
            //Surname
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[2].surname")).SendKeys("Matlou");

            //ID Number
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[2].id-number")).SendKeys("7001113329081");

            //Cellphone
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[2].contact-number")).SendKeys("0679774522");


            //select relationship
            Delay(2);
            //Cover Amount 
            SlideBar("Child");

            TakeScreenshot(_driver, $@"{_screenShotFolder}\validations\", "MaxChildValidation");

            //click next 
            Delay(2);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/button")).Click();

            Delay(2);


            //Max validation
            var Errormessage = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper]/article/form/section[5]/div[4]/div[1]/label")).Text;



            if ((Errormessage) != ("Cover Amount of R20 000 at R28 per month"))
            {
                results = "Passed";
            }
            else
            {
                results = "Failed";
            }



            return results;

        }

        [Category("ExtendedMaxAge")]
        public string ExtendedMaxAge()
        {


            string results = "";


            //Add Life
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/button")).Click();

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
            _driver.FindElement(By.Name("/cover-details[4].name")).SendKeys("Zondi");
            //Surname
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[4].surname")).SendKeys("Zwane");

            //ID Number
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[4].id-number")).SendKeys("3011148884087");

            //Cellphone
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[4].contact-number")).SendKeys("0670974589");


            //Cover Amount 

            SlideBar("Extended");



            TakeScreenshot(_driver, $@"{_screenShotFolder}\Failed_Scenarios\", "ExtendedMaxvalidation");

            //click next 
            Delay(2);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/button")).Click();

            Delay(2);


            //Max validation
            var Errormessage = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[7]/div[4]/div[1]/label")).Text;


            if ((Errormessage) == ("Cover is only available for persons up to 85 years of age"))
            {
                results = "Passed";
            }
            else
            {
                results = "Failed";
            }

            return results;

        }

        private string SpouseMaxMin()
        {
            String results;
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
            _driver.FindElement(By.Name("/cover-details[1].id-number")).SendKeys("0609162644080");

            //Cellphone
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[1].contact-number")).SendKeys("0679774589");


            //Cover Amount 

            SlideBar("Spouse");



            String minimumValidationMsg = _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/form/section[4]/div[4]/div[1]/label")).Text;
            if (minimumValidationMsg.Contains("Cover is only available for spouses from 18 to 64 years of age"))
            {
                TakeScreenshot(_driver, $@"{_screenShotFolder}\Validations\", "Spouse_Min_Age");


            }

            //click on id field
            Actions builder = new Actions(_driver);
            builder.MoveToElement(_driver.FindElement(By.Name("/cover-details[1].id-number")))
            .Click().Build().Perform();

            //clear keys
            IWebElement clear = _driver.FindElement(By.Name("/cover-details[1].id-number"));
            clear.SendKeys(Keys.Control + "a");
            clear.SendKeys(Keys.Delete);
            Delay(3);



            //Send new keys
            IWebElement element = _driver.FindElement(By.Name("/cover-details[1].id-number"));
            Delay(3);
            element.SendKeys("3403148170080");
            Delay(2);


            String maximumValidationMsg = _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/form/section[4]/div[4]/div[1]/label")).Text;
            if (maximumValidationMsg.Contains("Cover is only available for spouses from 18 to 64 years of age"))
            {
                TakeScreenshot(_driver, $@"{_screenShotFolder}\Failed_Scenarios\", "Spouse_Max_Age");
            }

            if (maximumValidationMsg.Contains("Cover is only available for spouses from 18 to 64 years of age") && minimumValidationMsg.Contains("Cover is only available for spouses from 18 to 64 years of age"))
            {
                results = "Passed";
            }
            else
            {
                Delay(1);
                results = "Failed";
                TakeScreenshot(_driver, $@"{_screenShotFolder}\Failed_Scenarios\", "Spouse_Max_Min_Validations");
            }


            //Cover Amount 

            SlideBar("Spouse");
        
            return results;

        }

        public string createNewClient()
        {
            string results;
            //get policy holder data
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
                TakeScreenshot(_driver, $@"{_screenShotFolder}\validations\", "MinMainLifeValidation");
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
                
                TakeScreenshot(_driver, $@"{_screenShotFolder}\validations\", "MaxMainLifeValidation");
            }

            if ((Errormessage2) == ("Must not be older than 74 years of age.") && (Errormessage) == ("Must be at least 18 years old."))
            {
                results = "Passed";
            }
            else
            {
                results = "Failed";
                TakeScreenshot(_driver, $@"{_screenShotFolder}\failed_scenarios\", "MaxMainLifeValidation");

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


            return results;

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


        [TearDown]
        public void closeBrowser()
        {
            base.DisconnectBrowser();
        }

    }
}
