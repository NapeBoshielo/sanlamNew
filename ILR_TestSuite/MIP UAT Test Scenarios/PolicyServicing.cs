using NUnit.Framework;

using static ILR_TestSuite.TestBase;

using OpenQA.Selenium;

using OpenQA.Selenium.Interactions;

using System;

using System.Collections.Generic;
using ILR_TestSuite;
using System.Linq;
using OpenQA.Selenium.Support.UI;
using System.Data.OleDb;
using Actions = OpenQA.Selenium.Interactions.Actions;
using System.Data;


namespace PolicyServicing

{

    [TestFixture]

    public class PolicyServicing : TestBase
    {

        //Policy-Servicing

        private string sheet;



        [SetUp]

        public void startBrowser()

        {



            _driver = base.SiteConnection();
            sheet = "Policy-Servicing";

        }



        [Test, Order(1)]

        public void PolicyServicingTestSuite()

        {




            Delay(2);
            //Click on Miain
            _driver.FindElement(By.Name("CBWeb")).Click();
            // Delay(2);
             //ChangeBeneficiary();
            // CancelPolicy();
            DecreaseSumAssured();


            Delay(20);

        }
        private void DecreaseSumAssured()
        {
            try
            {
                string contRef = base.GetPolicyNoFromExcell(sheet, "DecreaseSumAssured");

                string results = "";
            
                var currentSumAssured = "";
                var currentPremium = "";
                var newPremium = "";
                var commDate = "";

                policySearch(contRef);

                //Get the Commencement date from contract summary screen
                commDate = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr[6]/td[2]")).Text;
                //Scroll Down
                Delay(2);

                IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;
                

                Delay(4);

                //Get Current premium
                currentPremium = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv9']/table/tbody/tr[2]/td[2]")).Text;

                Delay(4);

                //Select the component
                _driver.FindElement(By. Name("fccComponentDescription1")).Click();




                Delay(4);
                //Get The current Sum Assured for the life assured
                currentSumAssured = _driver.FindElement(By.XPath("//*[@id='frmCbmcc']/tbody/tr[8]/td[4]")).Text;



                IWebElement policyOptionElement = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/div[3]/table/tbody/tr/td/div/div[3]"));

                //Creating object of an Actions class
                Actions action = new Actions(_driver);

                //Performing the mouse hover action on the target element.
                action.MoveToElement(policyOptionElement).Perform();

                Delay(5);

                _driver.FindElement(By.XPath("//div[3]/a/img")).Click();

                Delay(4);
                _driver.FindElement(By.Name("frmCCStartDate")).Clear();
                Delay(2);
                _driver.FindElement(By.Name("frmCCStartDate")).SendKeys(commDate);

                var newSumAssured ="";
                //Do a  downgrade on current sum assured by 5000
                if (Convert.ToInt32(currentSumAssured) > 10000 || Convert.ToInt32(currentSumAssured) == 10000)
                {
                   newSumAssured = (Convert.ToInt32(currentSumAssured) - 10000).ToString();
                }
                else
                {
                    newSumAssured = (5000).ToString();
                }

                SelectElement oSelect = new SelectElement(_driver.FindElement(By.Name("frmSPAmount")));

                oSelect.SelectByValue(newSumAssured);
               

                Delay(4);
                _driver.FindElement(By.Name("btncbmcc13")).Click();
                Delay(4);
                _driver.FindElement(By.Name("btncbmcc17")).Click();
                //Calculate age based on IdNo
                var idElement = _driver.FindElement(By.XPath("//*[@id='frmCbmcc']/tbody/tr[9]/td[2]")).Text;
                var idNo = (idElement.Split(" ")[idElement.Split(" ").Length - 1]).ToString();
                var birthYear = idNo.Substring(1, 2);
                birthYear = "19" + birthYear;
                var age = (DateTime.Now.Year - Convert.ToInt32(birthYear)).ToString();

                var premuimfromRateTable = base.getPremuimFromRateTable(age, "ML", newSumAssured, "Serenity_Premium");

                Delay(4);
                _driver.FindElement(By.Name("btncbmcc23")).Click();
                Delay(6);

                //Get the new Premium
                //js.ExecuteScript("window.scrollBy(0,1000)", "");
                newPremium = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv5']/table/tbody/tr[2]/td[7]")).Text;

                //Do Age validation
                string dateInput = "9905208629080";
                var splitted = dateInput.Substring(0,2);
              

                //var rateTablePremium = base.getPremuimFromRateTable();


                Delay(4);
                //Go Back to contract summary
                _driver.FindElement(By.Name("PF_User_Menu")).Click();
                _driver.FindElement(By.Name("cb_User_cbmct")).Click();
                Delay(4);
                
                if (premuimfromRateTable == Convert.ToDecimal(newPremium))
                {
                    results = "Passed";
                }
                else
                {
                    results = "Failed";
                }
                base.writeResultsToExcell(results, sheet, "DecreaseSumAssured");


            }
            catch (Exception)
            {

                throw;
            }
        }
        private void ChangeBeneficiary()
        {
            try
            {
                string contRef = base.GetPolicyNoFromExcell(sheet, "ChangeBeneficiary");

                string results = "";
                policySearch(contRef);

                
                IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;
                js.ExecuteScript("window.scrollBy(0,500)", "");
                Delay(2);

                for (int i = 2; i < 23; i++)
                {
                    var txt = _driver.FindElement(By.XPath($"//*[@id='CntContentsDiv11']/table/tbody/tr[{i.ToString()}]/td[1]/span")).Text;
                    if(txt == "Beneficiary")
                    {
                        _driver.FindElement(By.XPath($"//*[@id='CntContentsDiv11']/table/tbody/tr[{i.ToString()}]/td[2]/a")).Click();
                        break;
                    }
                   
                }
                string title ="", name="", surname="", dob="", gender="", id_no="";

                //Extract data from excell
                using (OleDbConnection conn = new OleDbConnection(base._connString))
                {
                    try
                    {

                        // Open connection
                        conn.Open();
                        string cmdQuery = "SELECT * FROM [ChangeBeneData$]";

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


                            title = ((System.Data.DataRowView)row).Row.ItemArray[0].ToString();
                            name = ((System.Data.DataRowView)row).Row.ItemArray[1].ToString();
                            surname = ((System.Data.DataRowView)row).Row.ItemArray[2].ToString();
                            dob = ((System.Data.DataRowView)row).Row.ItemArray[3].ToString();
                            gender = ((System.Data.DataRowView)row).Row.ItemArray[4].ToString();
                            id_no = ((System.Data.DataRowView)row).Row.ItemArray[5].ToString();


                        }




                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    finally
                    {
                        conn.Close();
                        conn.Dispose();
                    }
                }
                

                Delay(2);
                //Press on change button
                _driver.FindElement(By.Name("btnChangePerson")).Click();

                Delay(2);

              

                var value = "";
                switch (title)
                {
                    case "Mr":
                        value = "er_AcPerTitleMr";
                        break;
                    case "Mrs":
                        value = "er_AcPerTitleMrs";
                        break;

                    case "Ms":
                        value = "er_AcPerTitleMs";
                        break;

                    case "Prf":
                        value = "er_AcPerTitlePrf";
                        break;
                    case "Dr":
                        value = "er_AcPerTitleDoc";
                        break;

                    case "Adm":
                        value = "er_AcPerTitleADM";
                        break;

                    case "Miss":
                        value = "er_AcPerTitleMiss";
                        break;

                    default:
                        break;
                }

                SelectElement oSelect = new SelectElement(_driver.FindElement(By.Name("fcTitle")));
                //Select title
                oSelect.SelectByValue(value);
                Delay(2);
                //Insert Name
                _driver.FindElement(By.Name("fcFirstName")).Clear();
                _driver.FindElement(By.Name("fcFirstName")).SendKeys(name);

                Delay(2);

                //Insert surname
                _driver.FindElement(By.Name("fcLastName")).Clear();
                _driver.FindElement(By.Name("fcLastName")).SendKeys(surname);
                Delay(2);

                //Insert Date of birth
                var oldDOB = _driver.FindElement(By.Name("fcDateOfBirth")).Text;
                _driver.FindElement(By.Name("fcDateOfBirth")).Clear();
                _driver.FindElement(By.Name("fcDateOfBirth")).SendKeys(dob);
                Delay(2);

                switch (gender)
                {
                    case ("Male"):
                        value = "er_AcPerGenMal";
                        break;
                    case ("Female"):
                        value = "er_AcPerGenFem";
                            break;
                    default:
                        value = "er_AcPerGenGna";
                        break;
                }
                oSelect = new SelectElement(_driver.FindElement(By.Name("fcGender")));
                //Select gender
                oSelect.SelectByValue(value);
                Delay(2);
                _driver.FindElement(By.Name("btnSubmit")).Click();
                Delay(2);

                js.ExecuteScript("window.scrollBy(0,4000)", "");
                Delay(2);
                _driver.FindElement(By.Name("fccIdentityType2")).Click();
                Delay(2);

                _driver.FindElement(By.Name("btnagmpi2")).Click();
                Delay(2);

                //Change ID number
                var oldID = _driver.FindElement(By.Name("frmIdNumber")).Text;
                _driver.FindElement(By.Name("frmIdNumber")).Clear();
                _driver.FindElement(By.Name("frmIdNumber")).SendKeys(id_no);
                //Enter issue date
                var todayDt = DateTime.Today;

                _driver.FindElement(By.Name("frmIssueDate")).SendKeys($"{todayDt.Year}/{todayDt.Month}/{todayDt.Day}");
                Delay(2);
                _driver.FindElement(By.Name("btnagmpi0")).Click();
                Delay(2);


                //Vaidation based on ID
                if (oldDOB != dob && oldID != id_no)
                {
                    results = "Passed";
                }
                else
                {
                    results = "Failed";
                }
              

                base.writeResultsToExcell(results, sheet, "ChangeBeneficiary");

            }
            catch (Exception ex)
            {
                DisconnectBrowser();

                throw ex;
            }
        }


        [Category("Cancel Policy")]
        private void CancelPolicy()
        {

            try

            {

                string contRef = base.GetPolicyNoFromExcell(sheet, "CancelPolicy");

                string results = "";

                string date = DateTime.Today.ToString("g");


                policySearch(contRef);
            
                Delay(3);
                //Hover on policy options
                IWebElement policyOptionElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));

                //Creating object of an Actions class
                Actions action = new Actions(_driver);

                //Performing the mouse hover action on the target element.
                action.MoveToElement(policyOptionElement).Perform();

                Delay(5);
                //Click on Cancel

                _driver.FindElement(By.XPath("//table[@id='m0t0']/tbody/tr/td/div/div[3]/a/img")).Click();
                Delay(5);
                SelectElement oSelect = new SelectElement(_driver.FindElement(By.Name("frmCancelReason")));

                oSelect.SelectByValue("Cancelled by external service");
                Delay(4);
                //cancel
                _driver.FindElement(By.Name("btnSubmit")).Click();
                Delay(2);



                // Switch the control of 'driver' to the Alert from main Window
                IAlert simpleAlert1 = _driver.SwitchTo().Alert();

                // '.Accept()' is used to accept the alert '(click on the Ok button)'
                simpleAlert1.Accept();

                Delay(2);

                //2000175333.8
                _driver.FindElement(By.Name("2000175333.8_expand")).Click();
                Delay(2);

                _driver.FindElement(By.Name("ICF_00000648")).Click();
                Delay(2);
                var movementType = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv1']/center/table/tbody/tr[2]/td/center/table/tbody/tr[2]/td[1]")).Text;



      
                    //Assert.Pass("The policy was succesfully cancelled");
                    results = "Passed";

               
                base.writeResultsToExcell(results, sheet, "CancelPolicy");

            }

            catch (Exception ex)

            {

                DisconnectBrowser();

                throw ex;

            }

        }


        [Category("Policy Search")]
        public void policySearch( string contractRef = "")
        {
          
            IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;

            
            //Click on contract search 
            _driver.FindElement(By.Name("alf-ICF8_00000222")).Click();
            Delay(2);


            //Click on product
            _driver.FindElement(By.Name("frmProductCode")).Click();
            Delay(2);

            //Type in contract ref if there is any 
      
              _driver.FindElement(By.Name("frmContractReference")).SendKeys(contractRef);


          
            Delay(4);

            //Click on Search Icon 
            _driver.FindElement(By.Name("btncbcts0")).Click();
            Delay(2);
            _driver.FindElement(By.XPath("//*[@id='AppArea']/table[2]/tbody/tr[2]/td[1]/a")).Click();

        }

        [Category("LogInValidation")]

        public void LogInValidation()

        {

            try

            {
                //Arrange
                var expectedWelcomePage = "Welcome to Sanlam ARL Demo for the Web";
               

                //Action

                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/table/tbody/tr[2]/td[3]/center[1]/b"));

               TakeScreenshot(_driver, $@"{_screenShotFolder}\LogInValidation\", "Landing Page Validation");

                //find the result
                string actualResult = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/table/tbody/tr[2]/td[3]/center[1]/b")).Text;

                Assert.IsTrue(expectedWelcomePage.Equals(actualResult, StringComparison.CurrentCultureIgnoreCase));


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
