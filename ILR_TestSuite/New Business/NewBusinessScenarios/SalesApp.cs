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
                Delay(15);
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
                            var result = PositiveTestProcess(Scenario_ID);
                            writeResultsToExcell(results, "Scenarios",Scenario_ID);

                        }
                    }
                    finally
                    {
                        conn.Close();
                        conn.Dispose();
                    }
                }

            }
  
        public Dictionary<string, string> createNewClient(string scenario_ID)
        {

            string results = "";
            //get policy holder data
            var policyHolderData = getPolicyHolderDetails(scenario_ID);
            _driver.SwitchTo().ActiveElement();
            _driver.FindElement(By.XPath("//*[@id='___gatsby']"));
            Delay(20);
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

            return policyHolderData;
        }
        public string PositiveTestProcess(string scenario_ID)
        {
            var upload_file = "C:/Users/G992127/Documents/GitHub/ILR_TestSuite/ILR_TestSuite/New Business/upload/download.jpg";
            var plaData = createNewClient(scenario_ID);
            string results="";
            var policyplayers = getRolePlayers(scenario_ID);
            List<string> keys = new List<string>();
            keys.Add("PolicyHolder_Details");
            keys.Add("spouse");
            keys.Add("Children");
            keys.Add("Parents");
            keys.Add("Extended");

            var beneficiaries = policyplayers["Beneficiaries"];

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
           //Add Provided LAs
            var lifeAsuredCounter = 0;
            var label = 1;
            var section = 3;
            foreach (var key in keys) 
            {

                foreach (var item in policyplayers[key])
                {
                     if(item.Count > 0)
                    {
                        if (key == "PolicyHolder_Details")
                        {
                            if (item["Covered"] == "Yes")
                            {

                                //add main life
                                Delay(1);
                                _driver.FindElement(By.XPath($"//*[@id='gatsby-focus-wrapper']/article/form/section[{section}]/div[2]/div[1]/div/label[{label}]")).Click();
                                //Cover Amount
                                SlideBar("Myself");
                                section++;
                                lifeAsuredCounter++;
                                break;
                            }
                        }
                        //click Add 
                        Delay(2);
                        _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/button")).Click();

                        //select relationship
                        Delay(2);
                        _driver.FindElement(By.XPath($"//*[@id='gatsby-focus-wrapper']/article/form/section[{section}]/div[2]/div[1]/div/label[{label}]")).Click();
                        if (key == "Extended")
                        {
                            //Extended Relationship Type

                            IWebElement RelationshipType = _driver.FindElement(By.Name($"/cover-details[{lifeAsuredCounter}].relationship-extended-type"));
                            RelationshipType.SendKeys(item["Extended_RelationshipType"]);
                            RelationshipType.SendKeys(Keys.ArrowDown);
                            RelationshipType.SendKeys(Keys.Enter);
                        }
                        //FirstName
                        Delay(1);
                        _driver.FindElement(By.Name($"/cover-details[{lifeAsuredCounter}].name")).SendKeys(item["First_name"]);
                        //Surname
                        Delay(2);
                        _driver.FindElement(By.Name($"/cover-details[{lifeAsuredCounter}].surname")).SendKeys(item["Surname"]);
                        //ID Number
                        Delay(1);
                        _driver.FindElement(By.Name($"/cover-details[{lifeAsuredCounter}].id-number")).SendKeys(item["ID_number"]);
                        //Cellphone
                        Delay(2);
                        _driver.FindElement(By.Name($"/cover-details[{lifeAsuredCounter}].contact-number")).SendKeys(item["Cellphone"]);

                        section++;
                        lifeAsuredCounter++;
                    }
                    else
                    {
                        break;
                    }
                }
            
                label++;
                
            }


            Delay(1);
   
            //Click next
            Delay(3);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();


            var beneCounter = 0;
            //payment reciever(Beneficiary)
            foreach (var item in beneficiaries)
            {
                //click relationship 
                Delay(1);
                _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/div/section/div[3]/div/div[1]/div/label[1]")).Click();

                //FirstName
                Delay(1);
                _driver.FindElement(By.Name($"/funeral-beneficiaries[{beneCounter}].name")).SendKeys(item["First_name"]);
                //Surname
                Delay(1);
                _driver.FindElement(By.Name($"/funeral-beneficiaries[{beneCounter}].surname")).SendKeys(item["Surname"]);

                //ID Number
                Delay(1);
                _driver.FindElement(By.Name($"/funeral-beneficiaries[{beneCounter}].id-number")).SendKeys(item["ID_number"]);

                //Cellphone
                Delay(1);
                _driver.FindElement(By.Name($"/funeral-beneficiaries[{beneCounter}].contact-number")).SendKeys(item["Cellphone"]);

                //Percentage
                IWebElement sliderbar5 = _driver.FindElement(By.ClassName("slider"));
                int widthslider5 = sliderbar5.Size.Width;
                Delay(1);
                IWebElement slider5 = _driver.FindElement(By.ClassName("slider"));
                Actions slideraction5 = new Actions(_driver);
                slideraction5.ClickAndHold(slider5);
                slideraction5.MoveByOffset(260, 0).Build().Perform();
                beneCounter++;

            }


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
            string bank = plaData["Bank"];

            //policy payer
            Delay(1);
            _driver.FindElement(By.Name("/same-as-fna")).Click();


            //bank details
            //Bank Selction
      
            SelectElement oSelect1 = new SelectElement(_driver.FindElement(By.Name("/bank-select")));
            oSelect1.SelectByValue(bank);

            //Account Number
            Delay(1);
            _driver.FindElement(By.Name("/account-number")).SendKeys(plaData["Account_Number"]);


            //Account Type
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[1]/div[2]/div[4]/div/label[2]")).Click();


            ///debit - order - date / debit - order - date
            SelectElement oSelect = new SelectElement(_driver.FindElement(By.Name("/debit-order-date")));
            oSelect.SelectByValue(plaData["Debit_Order_Day"]);

            //salarypaydate
            Delay(1);
            _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/form/section[1]/div[2]/div[6]/input")).SendKeys(plaData["Salary_Date"]);

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
            _driver.FindElement(By.Name("/physical-address-building")).SendKeys(plaData["Building"]);
            //Street
            Delay(1);
            _driver.FindElement(By.Name("/physical-address-street")).SendKeys(plaData["Street"]);

            //Town
            Delay(1);
            _driver.FindElement(By.Name("/physical-address-town")).SendKeys(plaData["City"]);

            //Suburb
            Delay(1);
            _driver.FindElement(By.Name("/physical-address-suburb")).SendKeys(plaData["Suburb"]);

            //CodeField 
            _driver.FindElement(By.Name("/physical-address-code")).SendKeys(plaData["Code"]);



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
            _driver.FindElement(By.Name("/call-reference-number")).SendKeys(plaData["Call_Ref_Number"]);

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
            _driver.FindElement(By.Id("/identification")).SendKeys(upload_file);

            //upload2
            Delay(2);
            _driver.FindElement(By.Id("/q-link")).SendKeys(upload_file);
            //upload3
            Delay(1);
            _driver.FindElement(By.Id("/proof-of-income")).SendKeys(upload_file);


            //click next
            Delay(4);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div/a[2]")).Click();


            //Card number
            Delay(4);
            _driver.FindElement(By.Id("/card-number")).SendKeys(plaData["Card_Number"]);

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
