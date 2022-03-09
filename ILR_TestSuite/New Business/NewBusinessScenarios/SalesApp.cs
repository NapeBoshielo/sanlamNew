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
            Delay(40);
           // string scenario_id;
        // createNewClient();
           // writeResultsToExcell( "Scenarios", "MaxMinAgeMainLife");

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


                    //addMainLife();
                    foreach (var row in ds.Tables[0].DefaultView)
                    {
                        var Scenario_ID = ((System.Data.DataRowView)row).Row.ItemArray[0].ToString();
                        PositiveTestProcess(Scenario_ID);
                        //switch (func)
                        //{
                        //    case "MaxMinAgeSpouse":
                        //        SpouseMaxMin();
                        //    //  writeResultsToExcell(results, sheet, func);
                        //        break;
                        //    case "MaxMinAgeChild":
                        //        ChildMaxAge();
                        //       // writeResultsToExcell(results, sheet, func);
                        //        break;
                        //    case "MaxMinAgeParent":
                        //       MaxMinAgeParent();
                        //  //  writeResultsToExcell(results, sheet, func);
                        //        break;
                        //    case "MaxMinAgeExtended":
                        //         ExtendedMaxAge();
                        //       // writeResultsToExcell(results, sheet, func);
                        //        break;
                        //    case "PositiveTest":
                        //   // writeResultsToExcell(results, sheet, func);
                        //    break;
                        //default:
                        //        break;
                        //}


                    }
                    



                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }





        }
        private void MaxMinAgeParent()
        {
            String results;
            //click Add 
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/button")).Click();


            //select relationship
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[6]/div[2]/div[1]/div/label[4]")).Click();

            //FirstName
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[3].name")).SendKeys("Sizwe");
            //Surname
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[3].surname")).SendKeys("Ngamola");

            //ID Number
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[3].id-number")).SendKeys("6902075770082");

            //Cellphone
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[3].contact-number")).SendKeys("0820678899");


            //Cover Amount 

            SlideBar("Parent");



            String minimumValidation = _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/form/section[6]/div[4]/div[1]/label")).Text;
           // if (minimumValidation.Contains("Cover is only available for parents from 26 to 85 years of age"))
           // {
           //     TakeScreenshot(_driver, $@"{_screenShotFolder}\validations\", "Parent_Min_Age");


          //  }

            //click on id field
            Actions builder = new Actions(_driver);
            builder.MoveToElement(_driver.FindElement(By.Name("/cover-details[3].id-number")))
            .Click().Build().Perform();

            //clear keys
            IWebElement clear = _driver.FindElement(By.Name("/cover-details[3].id-number"));
            clear.SendKeys(Keys.Control + "a");
            clear.SendKeys(Keys.Delete);
            Delay(3);



            //Send new keys
            IWebElement element = _driver.FindElement(By.Name("/cover-details[3].id-number"));
            Delay(3);
            element.SendKeys("6806081846085");
            Delay(2);


            String maximumValidation = _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/form/section[6]/div[4]/div[1]/label")).Text;
          //  if (maximumValidation.Contains("Cover is only available for parents from 26 to 85 years of age"))
          //  {
           //     TakeScreenshot(_driver, $@"{_screenShotFolder}\validations\", "Parent_Max_Age");
           // }
        //
           // if (maximumValidation.Contains("Cover is only available for parents from 26 to 85 years of age") && minimumValidation.Contains("Cover is only available for parents from 26 to 85 years of age"))
          //  {
           ///     results = "Passed";
           // }
           // else
           // {

              //  TakeScreenshot(_driver, $@"{_screenShotFolder}\Failed_Scenarios\", "Parent_Max_Min_Validations");

              //  results = "Failed";
           // }


            //Cover Amount 

            //SlideBar("Parent");

            //return results;

        }

        [Category("Add Main Life")]
        public void addMainLife()
        {

            ////click tickbox product
            //Delay(1);
            //_driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/section/div[1]/div[3]/button/span")).Click();

            ////click tickbox product
            //Delay(1);
            //_driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/div/div[2]/label/div")).Click();


            ////click next
            //Delay(1);
            //_driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a")).Click();



            ////click on No
            //Delay(1);
            //_driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[1]/div/div[2]/div/div/label[2]")).Click();


            ////click on 5%
            //Delay(1);
            //_driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[2]/div/div[2]/div/div/label[1]")).Click();




            ////add main life
            //Delay(1);
            //_driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[3]/div[2]/div[1]/div/label[1]")).Click();
            


            ////Cover Amount

            //SlideBar("Myself");
        }

        [Category("ChildMaxAge")]
        public void ChildMaxAge()
        {


            string results = "";

            ////Add life 
            //Delay(1);
            //_driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/button")).Click();

            ////select relationship
            //Delay(1);

            //// Locate the element “child” by By.xpath. 
            //IWebElement Child = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[5]/div[2]/div[1]/div/label[3]"));

            //// Create an object of Actions class and pass reference variable driver as a parameter to its constructor. 
            //Actions actions = new Actions(_driver);
            //actions.Click(Child);
            //actions.Build().Perform();


            ////FirstName
            //Delay(1);
            //_driver.FindElement(By.Name("/cover-details[2].name")).SendKeys("Zodwa");
            ////Surname
            //Delay(1);
            //_driver.FindElement(By.Name("/cover-details[2].surname")).SendKeys("Matlou");

            ////ID Number
            //Delay(1);
            //_driver.FindElement(By.Name("/cover-details[2].id-number")).SendKeys("1009248900086");   

            ////Cellphone
            //Delay(1);
            //_driver.FindElement(By.Name("/cover-details[2].contact-number")).SendKeys("0679774522");


            //select relationship
            Delay(2);


            //Cover Amount 
            //SlideBar("Child");

            


            Delay(2);


            //Max validation
            var Errormessage = _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/form/section[5]/div[4]/div[1]/label")).Text;
            

         /*   if ((Errormessage) == ("Cover is only available for children up to 25 years of age"))
            {

                TakeScreenshot(_driver, $@"{_screenShotFolder}\validations\", "MaxChildValidation");
                results = "Passed";
            }
            else
            {
                TakeScreenshot(_driver, $@"{_screenShotFolder}\failedscenarios\", "MaxChildValidation");

                results = "Failed";
            }



            return results;
         */
        }

        [Category("ExtendedMaxAge")]
        public void ExtendedMaxAge()
        {


            string results = "";


            //Add Life
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/button")).Click();

            //select relationship
            Delay(1);




            //Extended

            /// Locate the element “child” by By.xpath. 
            IWebElement extended = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[7]/div[2]/div[1]/div/label[5]"));

           

            // Create an object of Actions class and pass reference variable driver as a parameter to its constructor. 
            Actions actions = new Actions(_driver);
            actions.Click(extended);
            actions.Build().Perform();


            //Extended Relationship Type

            IWebElement RelationshipType = _driver.FindElement(By.Name("/cover-details[4].relationship-extended-type"));
            RelationshipType.SendKeys("uncle");
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
            _driver.FindElement(By.Name("/cover-details[4].id-number")).SendKeys("8803205554081");

            //Cellphone
            Delay(1);
            _driver.FindElement(By.Name("/cover-details[4].contact-number")).SendKeys("0670974589");


            //Cover Amount 

            SlideBar("Extended");




            Delay(2);


            //Max validation
            var Errormessage = _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/form/section[7]/div[4]/div[1]/label")).Text;


            /* if ((Errormessage) == ("Cover is only available for persons up to 85 years of age"))
             {

                 TakeScreenshot(_driver, $@"{_screenShotFolder}\validations\", "ExtendedMaxvalidation");

                 results = "Passed";
             }
             else
             {
                 results = "Failed";

                 TakeScreenshot(_driver, $@"{_screenShotFolder}\Failedscenarios\", "ExtendedMaxvalidation");
             }

           return results;
   */
        }

        private void SpouseMaxMin()
        {
            String results;
            ////click Add 
            //Delay(1);
            //_driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/button")).Click();



            ////select relationship
            //Delay(1);
            //_driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[4]/div[2]/div[1]/div/label[2]")).Click();

            ////FirstName
            //Delay(1);
            //_driver.FindElement(By.Name("/cover-details[1].name")).SendKeys("Lindokuhle");
            ////Surname
            //Delay(1);
            //_driver.FindElement(By.Name("/cover-details[1].surname")).SendKeys("Matlou");

            ////ID Number
            //Delay(1);
            //_driver.FindElement(By.Name("/cover-details[1].id-number")).SendKeys("8202075420087");

            ////Cellphone
            //Delay(1);
            //_driver.FindElement(By.Name("/cover-details[1].contact-number")).SendKeys("0679774589");


            //Cover Amount 

           // SlideBar("Spouse");



          String minimumValidationMsg = _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/form/section[4]/div[4]/div[1]/label")).Text;
          ///  if (minimumValidationMsg.Contains("Cover is only available for spouses from 18 to 64 years of age"))
           // {
           //     TakeScreenshot(_driver, $@"{_screenShotFolder}\Validations\", "Spouse_Min_Age");


      // }

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
            element.SendKeys("9002143493085");
            Delay(2);


          /*  String maximumValidationMsg = _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/form/section[4]/div[4]/div[1]/label")).Text;
            if (maximumValidationMsg.Contains("Cover is only available for spouses from 18 to 64 years of age"))
            {
                TakeScreenshot(_driver, $@"{_screenShotFolder}\validations\", "Spouse_Max_Age");
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
          */

            //Cover Amount 

            SlideBar("Spouse");
        
          // return results;
           
        }

        public void createNewClient(string scenario_ID)
        {
            string results = "";
            //get policy holder data
            var policyHolderData = getPolicyHolderDetails(scenario_ID);
            _driver.SwitchTo().ActiveElement();
            _driver.FindElement(By.XPath("//*[@id='___gatsby']"));
            Delay(1);
            IWebElement new_client = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/div[2]/div[1]/button"));
            new_client.Click();
            //  Actions action = new Actions(_driver);
            // action.MoveToElement(new_client).Perform()
            Delay(2);
            IWebElement town = _driver.FindElement(By.XPath("//*[@id='downshift-0-input']"));
            town.SendKeys(policyHolderData["Town"]);
            Delay(1);
            town.SendKeys(Keys.ArrowDown);
            Delay(1);
            town.SendKeys(Keys.Enter);
            Delay(4);
            IWebElement worksite = _driver.FindElement(By.XPath("//*[@id='downshift-1-input']"));
            worksite.SendKeys(policyHolderData["Worksite"]);
            Delay(1);
            worksite.SendKeys(Keys.ArrowDown);
            Delay(1);
            worksite.SendKeys(Keys.Enter);
            Delay(2);
            IWebElement employer = _driver.FindElement(By.XPath("//*[@id='downshift-2-input']"));
            employer.SendKeys(policyHolderData["Employment"]);
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
            _driver.FindElement(By.XPath("//*[@id='/name']")).SendKeys(policyHolderData["First_name"]);
            Delay(2);
            //maiden name
            _driver.FindElement(By.XPath("//*[@id='/maiden-surname']")).SendKeys(policyHolderData["Maiden_Surname"]);
            Delay(2);
            //Id 
            Delay(3);
            _driver.FindElement(By.XPath("//*[@id='/id-number']")).SendKeys(policyHolderData["ID_number"]);
            Delay(2);
            _driver.FindElement(By.XPath("//*[@id='/surname']")).SendKeys(policyHolderData["Surname"]);
            Delay(2);
            //Select ethicity
            IWebElement select = _driver.FindElement(By.XPath(" //*[@id='/ethnicity']"));
            SelectElement oselect = new SelectElement(select);
            oselect.SelectByValue(policyHolderData["Ethnicity"]);
            Delay(2);
            //Select Maratiel
            IWebElement selectstatus = _driver.FindElement(By.XPath("//*[@id='/marital-status']"));
            SelectElement cselect = new SelectElement(selectstatus);
            cselect.SelectByValue(policyHolderData["Marital_Status"]);
            Delay(2);
            //Enter contact number
            _driver.FindElement(By.XPath("//*[@id='/contact-number']")).SendKeys(policyHolderData["CellPhone_number"]);
            Delay(2);
            //Enter email
            _driver.FindElement(By.XPath("//*[@id='/email']")).SendKeys(policyHolderData["Email"]);
            Delay(2);
            //Enter gross monthly
            _driver.FindElement(By.XPath("//*[@id='/gross-monthly-income']")).SendKeys(policyHolderData["Gross"]);
            Delay(2);
            //Select employent type
            if (policyHolderData["Permanent"] == "Yes")
            {
                _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/section/form/div/div[16]/div/label[1]")).Click();
            }
            else
            {
                _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/section/form/div/div[16]/div/label[2]")).Click();
            }
            Delay(2);
            //Salary frequency
            switch (policyHolderData["Salary_frequency"])
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
            _driver.FindElement(By.Name("/total-salary-after-deductions")).SendKeys(policyHolderData["Net_Salary"]);
            //Additional income
            Delay(1);
            _driver.FindElement(By.Name("/additional-income")).SendKeys(policyHolderData["Additional_Income"]);
            //Existing Financial Cover
            Delay(1);
            _driver.FindElement(By.Name("/existing-financial-cover")).SendKeys(policyHolderData["Existing_FinancialCover"]);
            //School Fees
            Delay(1);
            _driver.FindElement(By.Name("/school-fees")).SendKeys(policyHolderData["School_Fees"]);
            //Food
            Delay(1);
            _driver.FindElement(By.Name("/food")).SendKeys(policyHolderData["Food"]);
            //Retail accounts
            Delay(1);
            _driver.FindElement(By.Name("/retail-accounts")).SendKeys(policyHolderData["Retail_accounts"]);
            //Cellphone
            Delay(1);
            _driver.FindElement(By.Name("/cellphone")).SendKeys(policyHolderData["Cellphone"]);
            //Debt
            Delay(1);
            _driver.FindElement(By.Name("/debt")).SendKeys(policyHolderData["Debt"]);
            // Mortgage / Rent
            Delay(1);
            _driver.FindElement(By.Name("/mortage-rent")).SendKeys(policyHolderData["Mortgage_Rent"]);
            //Transport
            Delay(1);
            _driver.FindElement(By.Name("/transport")).SendKeys(policyHolderData["Transport"]);
            //Entertainment / Other
            Delay(1);
            _driver.FindElement(By.Name("/entertainment-other")).SendKeys(policyHolderData["Entertainment_Other"]);
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
        public string PositiveTestProcess(string scenario_ID)
        {

            createNewClient(scenario_ID);
            string results="";
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


            var counter = 1;

            

         var policyplayers =   getRolePlayers(scenario_ID);
            //foreach (var keys in policyplayers) {
            //    if (policyplayers[keys] != "Extended") {
                
                
                
            //    }
            
            
            
            
            
            
            //}
                foreach (var player in policyplayers["spouse"]) {

                //click Add 
                Delay(1);
                _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/button")).Click();



                //select relationship
                Delay(1);
                _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[4]/div[2]/div[1]/div/label[2]")).Click();

                //FirstName
                Delay(1);
                _driver.FindElement(By.Name($"/cover-details[{counter}].name")).SendKeys(player["First_name"]);
                //Surname
                Delay(1);
                _driver.FindElement(By.Name($"/cover-details[{counter}].surname")).SendKeys(player["Surname"]);

                //ID Number
                Delay(1);
                _driver.FindElement(By.Name($"/cover-details[{counter}].id-number")).SendKeys(player["ID_number"]);

                //Cellphone
                Delay(1);
                _driver.FindElement(By.Name($"/cover-details[{counter}].contact-number")).SendKeys(player["Cellphone"]);
               
            }
            counter++;
            Delay(1);
            var counter1 = 1;
            foreach (var player in policyplayers["Children"])
            {
                Delay(1);
               
                //click Add 
                Delay(1);
                _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/button")).Click();



                //select relationship
                Delay(3);
                _driver.FindElement(By.XPath($"//*[@id='gatsby-focus-wrapper']/article/form/section[{counter1+4}]/div[2]/div[1]/div/label[3]")).Click();

                //*[@id="gatsby-focus-wrapper"]/article/form/section[5]/div[2]/div[1]/div/label[3]
                //FirstName
                Delay(1);
                _driver.FindElement(By.Name($"/cover-details[{counter}].name")).SendKeys(player["First_name"]);
                //Surname
                Delay(1);
                _driver.FindElement(By.Name($"/cover-details[{counter}].surname")).SendKeys(player["Surname"]);

                //ID Number
                Delay(1);
                _driver.FindElement(By.Name($"/cover-details[{counter}].id-number")).SendKeys(player["ID_number"]);

                //Cellphone
                Delay(1);
                _driver.FindElement(By.Name($"/cover-details[{counter}].contact-number")).SendKeys(player["Cellphone"]);
                counter1++;
            }
            counter++;
            foreach (var player in policyplayers["Parents"])
            {

                //click Add 
                Delay(1);
                _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/button")).Click();



                //select relationship
                Delay(1);
                _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[7]/div[2]/div[1]/div/label[4]")).Click();

                

                //FirstName
                Delay(1);
                _driver.FindElement(By.Name($"/cover-details[{counter}].name")).SendKeys(player["First_name"]);
                //Surname
                Delay(1);
                _driver.FindElement(By.Name($"/cover-details[{counter}].surname")).SendKeys(player["Surname"]);

                //ID Number
                Delay(1);
                _driver.FindElement(By.Name($"/cover-details[{counter}].id-number")).SendKeys(player["ID_number"]);

                //Cellphone
                Delay(1);
                _driver.FindElement(By.Name($"/cover-details[{counter}].contact-number")).SendKeys(player["Cellphone"]);

            }
            counter++;
            foreach (var player in policyplayers["Extended"])
            {

                //click Add 
                Delay(1);
                _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/button")).Click();



                //select relationship
                Delay(1);
                _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[4]/div[2]/div[1]/div/label[2]")).Click();


                // Create an object of Actions class and pass reference variable driver as a parameter to its constructor. 
                //Actions actions = new Actions(_driver);
                //actions.Click(extended);
                //actions.Build().Perform();


                //Extended Relationship Type

                IWebElement RelationshipType = _driver.FindElement(By.Name("/cover-details[4].relationship-extended-type"));
                RelationshipType.SendKeys("uncle");
                RelationshipType.SendKeys(Keys.ArrowDown);
                RelationshipType.SendKeys(Keys.Enter);

                //FirstName
                Delay(1);
                _driver.FindElement(By.Name($"/cover-details[{counter}].name")).SendKeys(player["First_name"]);
                //Surname
                Delay(1);
                _driver.FindElement(By.Name($"/cover-details[{counter}].surname")).SendKeys(player["Surname"]);

                //ID Number
                Delay(1);
                _driver.FindElement(By.Name($"/cover-details[{counter}].id-number")).SendKeys(player["ID_number"]);

                //Cellphone
                Delay(1);
                _driver.FindElement(By.Name($"/cover-details[{counter}].contact-number")).SendKeys(player["Cellphone"]);

            }
            counter++;
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

            //Percentage
            IWebElement sliderbar5 = _driver.FindElement(By.ClassName("slider"));
            int widthslider5 = sliderbar5.Size.Width;
            Delay(1);
            IWebElement slider5 = _driver.FindElement(By.ClassName("slider"));
            Actions slideraction5 = new Actions(_driver);
            slideraction5.ClickAndHold(slider5);
            slideraction5.MoveByOffset(260, 0).Build().Perform();


            //Click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();




            //Word of advice
            Delay(1);
            _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/form/section/div[2]/div/textarea")).SendKeys("Test");


            //Click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();

            //Click No
            Delay(3);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/section/div[2]/form/div[2]/div/label[2]")).Click();

            //*[@id="gatsby-focus-wrapper"]/article/section/div[2]/form/div[2]/div/label[2]
            //go to payment 
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div/a[2]")).Click();





            /////////Payment Details
            string Bank = "";

            //policy payer
            Delay(1);
            _driver.FindElement(By.Name("/same-as-fna")).Click();


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

            SelectElement oSelect1 = new SelectElement(_driver.FindElement(By.Name("/bank-select")));
            oSelect1.SelectByValue("FIRST NATIONAL BANK");

            //Account Number
            Delay(1);
            _driver.FindElement(By.Name("/account-number")).SendKeys("62429363625");


            //Account Type
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[1]/div[2]/div[4]/div/label[2]")).Click();



            //Bank Selction
            var value1 = "";
            switch (Bank)
            {
                case "1st":
                    value1 = "1";
                    break;
                case "15th":
                    value1 = "15";
                    break;

                case "20th":
                    value1 = "20";
                    break;

                case "25th":
                    value1 = "25";
                    break;
                case "30th":
                    value1 = "30";
                    break;

                case "Last day of the month":
                    value1 = "31";
                    break;

                default:
                    break;
            }

            ///debit - order - date / debit - order - date
            SelectElement oSelect = new SelectElement(_driver.FindElement(By.Name("/debit-order-date")));
            oSelect.SelectByValue("25");

            //salarypaydate
            Delay(1);
            _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/form/section[1]/div[2]/div[6]/input")).SendKeys("25");


            //click tickbox
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='/arrange-payment-gather-information-disclaimer']")).Click();

            //click yes
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[2]/section/div[1]/div/div/label[1]")).Click();

            //click yes
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[2]/section/div[2]/div/div/label[1]")).Click();


            //click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div/a")).Click();


            //click i uderstand
            Delay(1);
            IWebElement iagree = _driver.FindElement(By.XPath("/html/body/reach-portal/div/div/div/button"));
            iagree.Click();

            //click start
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/section/div[3]/button")).Click();

            //debicheck loading delay
            Delay(150);


            var Errormessage = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/section/div[2]/div[2]/div[2]/p")).Text;

            if (Errormessage == "DebiCheck accepted by customer")
            {

            

                Delay(1);
                //Click next
                _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div/a[2]")).Click();



            }
            else
            {
                
                //Click next
                _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div/button")).Click();



                Delay(1);
                SelectElement oSelect2 = new SelectElement(_driver.FindElement(By.Name("reason-for-skipping")));

                oSelect2.SelectByValue("DEBICHECK_KEEPS_FAILING");


                Delay(1);
                //click skip
                _driver.FindElement(By.XPath("/html/body/reach-portal/div/div/div/div/button[2]")).Click();


                Delay(1);
                //click skip
                _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div/a[2]")).Click();

            }


            //Physical Address

            //Building
            Delay(1);
            _driver.FindElement(By.Name("/physical-address-building")).SendKeys("27 Beacon Avenue");
            //Street
            Delay(1);
            _driver.FindElement(By.Name("/physical-address-street")).SendKeys("27 Beacon Avenue");

            //Town
            Delay(1);
            _driver.FindElement(By.Name("/physical-address-town")).SendKeys("Linbro Park");

            //Suburb
            Delay(1);
            _driver.FindElement(By.Name("/physical-address-suburb")).SendKeys("Sandton");

            //CodeField 
            _driver.FindElement(By.Name("/physical-address-code")).SendKeys("2090");



            ///click tickbox same-as-physical
            Delay(1);
            _driver.FindElement(By.Name("/same-as-physical")).Click();

            ///click tickbox 
            Delay(1);
            _driver.FindElement(By.Name("/policy-holder-signature-datetime")).Click();

            ///click tickbox 
            Delay(1);
            _driver.FindElement(By.Name("/premium-payer-signature-datetime")).Click();

            //reference no 
            Delay(1);
            _driver.FindElement(By.Name("/call-reference-number")).SendKeys("09876567");

            //click next 
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();

            
            ///click tickbox 
            Delay(1);
            _driver.FindElement(By.Name("/popia-consent-datetime")).Click();


            //click next 
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div/a[2]")).Click();



            Delay(2);
            //upload1
            _driver.FindElement(By.Id("/identification")).SendKeys("C:/Users/E697642/Documents/GitHub/ILR_TestSuite/ILR_TestSuite/New Business/upload/download.jpg");

            //upload2
            Delay(2);
            _driver.FindElement(By.Id("/q-link")).SendKeys("C:/Users/E697642/Documents/GitHub/ILR_TestSuite/ILR_TestSuite/New Business/upload/download.jpg");
            //upload3
            Delay(1);
            _driver.FindElement(By.Id("/proof-of-income")).SendKeys("C:/Users/E697642/Documents/GitHub/ILR_TestSuite/ILR_TestSuite/New Business/upload/download.jpg");


            //click next
            Delay(4);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div/a[2]")).Click();


            //Card number
            Delay(4);
            _driver.FindElement(By.Id("/card-number")).SendKeys("10000008");

            //next
            Delay(2);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();



            //next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div/a")).Click();


            Delay(15);


            //sync
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/nav/div/div/button")).Click();




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
