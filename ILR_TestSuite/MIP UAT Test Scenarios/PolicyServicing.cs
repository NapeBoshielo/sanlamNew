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
            // ChangeBeneficiary();
            // CancelPolicy();
            //IncreaseSumAssuredAge();
            //ReInstate();
            //IncreaseSumAssured();
            ChangeCollectionM();






            Delay(20);

        }
        private void IncreaseSumAssuredAge()
        {
            try
            {

                String test_url_3 = "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?";
                String test_url_4_title = "DateTime Picker";
                IJavaScriptExecutor js2 = (IJavaScriptExecutor)_driver;

                string contRef = base.GetPolicyNoFromExcell(sheet, "IncreaseSumAssuredAge");

                string results = "";

                string date = DateTime.Today.ToString("g");


                policySearch(contRef);

                Delay(3);

                //Click on user under component
                _driver.FindElement(By.Name("fccComponentDescription1")).Click();
                Delay(2);



                IWebElement policyOptionElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));


                //Creating object of an Actions class
                Actions action = new Actions(_driver);



                //Performing the mouse hover action on the target element.
                action.MoveToElement(policyOptionElement).Perform();

                //Click on options
                _driver.FindElement(By.XPath("//*[@id='m0t0']/tbody/tr[1]/td/div/div[3]/a/img")).Click();

                Delay(5);

                SelectElement oSelect4 = new SelectElement(_driver.FindElement(By.Name("frmSPAmount")));

                oSelect4.SelectByValue("50000");
                Delay(2);



                //Click on next
                _driver.FindElement(By.XPath("//*[@id='frmCbmcc']/tbody/tr[2]/td[2]/a/span/img")).Click();
                Delay(2);



                Assert.AreEqual(2, _driver.WindowHandles.Count);

                var newWindowHandle = _driver.WindowHandles[1];
                Assert.IsTrue(!string.IsNullOrEmpty(newWindowHandle));

                /* Assert.AreEqual(driver.SwitchTo().Window(newWindowHandle).Url, "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?"); */
                string expectedNewWindowTitle = test_url_4_title;
                Assert.AreEqual(_driver.SwitchTo().Window(newWindowHandle).Title, expectedNewWindowTitle);



                //Click on next
                _driver.FindElement(By.Id("aIncYear")).Click();
                Delay(2);

                //Click on next
                _driver.FindElement(By.Name("fcCal1")).Click();
                Delay(2);


                /* Return to the window with handle = 0 */
                _driver.SwitchTo().Window(_driver.WindowHandles[0]);
                Delay(4);



                //Click on next
                _driver.FindElement(By.Name("btncbmcc13")).Click();
                Delay(2);

                //Click on roleplayer
                _driver.FindElement(By.Name("frmRolePlayers")).Click();
                Delay(2);


                SelectElement oSelect = new SelectElement(_driver.FindElement(By.Name("frmRolePlayers")));

                oSelect.SelectByValue("966829447.488");

                //Click on next
                _driver.FindElement(By.Name("btncbmcc17")).Click();
                Delay(2);



                // Click on finish
                _driver.FindElement(By.Name("btncbmcc23")).Click();
                Delay(5);


                // Click on contract summa
                _driver.FindElement(By.Name("2000175333.8")).Click();
                Delay(5);




                IJavaScriptExecutor js3 = (IJavaScriptExecutor)_driver;
                js3.ExecuteScript("window.scrollTo(0, 250)");


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


                //*[@id="CntContentsDiv9"]/table/tbody/tr[5]/td[2]



                //Assert.Pass("The policy was succesfully cancelled");
                results = "Passed";


                base.writeResultsToExcell(results, sheet, "IncreaseSumAssured");

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

                js.ExecuteScript("window.scrollBy(0,5000)", "");
                Delay(2);
                _driver.FindElement(By.Name("fccIdentityType2")).Click();
                Delay(2);
                _driver.FindElement(By.Name("btnagmpi2")).Click();
                Delay(2);
                _driver.FindElement(By.Name("frmIdNumber")).Clear();
                _driver.FindElement(By.Name("frmIdNumber")).SendKeys(id_no);
                Delay(2);
                _driver.FindElement(By.Name("btnagmpi0")).Click();
                Delay(2);
                results = "Passed";
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
            
                Delay(5);



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
        [Category("ReInstate")]

        public void ReInstate()

        {

            try

            {

                String test_url_1 = "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?";
                String test_url_1_title = "MIP - Sanlam ARL - Warpspeed Lookup Window";
                IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;

                String test_url_3 = "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?";
                String test_url_4_title = "DateTime Picker";
                IJavaScriptExecutor js2 = (IJavaScriptExecutor)_driver;

                string contRef = base.GetPolicyNoFromExcell(sheet, "ReInstate");

                string results = "";

                string date = DateTime.Today.ToString("g");


                policySearch(contRef);

                Delay(5);


                //Hover on policy options
                IWebElement policyOptionElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));


                //Creating object of an Actions class
                Actions action = new Actions(_driver);


                //Performing the mouse hover action on the target element.
                action.MoveToElement(policyOptionElement).Perform();
                Delay(2);

                //Click on Cancel
                _driver.FindElement(By.XPath("//*[@id='m0t0']/tbody/tr[1]/td/div/div[3]/a/img")).Click();
                Delay(4);

                SelectElement oSelect1 = new SelectElement(_driver.FindElement(By.Name("frmCancelReason")));

                oSelect1.SelectByValue("Got Better Offer");
                Delay(4);

                //Click on calender
                _driver.FindElement(By.XPath("//*[@id='CntContentsDiv4']/table/tbody/tr[6]/td[2]/span/span/img")).Click();
                Delay(2);



                Assert.AreEqual(2, _driver.WindowHandles.Count);



                var newWindowHandle = _driver.WindowHandles[1];
                Assert.IsTrue(!string.IsNullOrEmpty(newWindowHandle));



                /* Assert.AreEqual(driver.SwitchTo().Window(newWindowHandle).Url, "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?"); */
                string expectedNewWindowTitle = test_url_4_title;
                Assert.AreEqual(_driver.SwitchTo().Window(newWindowHandle).Title, expectedNewWindowTitle);



                //Click on arrow
                _driver.FindElement(By.Id("aDecYear")).Click();
                Delay(2);




                //Click on Date
                _driver.FindElement(By.Name("fcCal1")).Click();
                Delay(2);




                /* Return to the window with handle = 0 */
                _driver.SwitchTo().Window(_driver.WindowHandles[0]);
                Delay(4);

                //cancel  
                _driver.FindElement(By.Name("btnSubmit")).Click();
                Delay(2);

                // Switch the control of 'driver' to the Alert from main Window
                IAlert simpleAlert1 = _driver.SwitchTo().Alert();

                // '.Text' is used to get the text from the Alert
                String alertText2 = simpleAlert1.Text;
                Console.WriteLine("Alert text is " + alertText2);

                // '.Accept()' is used to accept the alert '(click on the Ok button)'
                simpleAlert1.Accept();

                Delay(2);


                //Contract Status validation 

                var StatusNotTakenUp = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr[2]/td[2]/u/font")).Text;



                IWebElement policyOptionElement3 = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));


                //Creating object of an Actions class
                Actions action2 = new Actions(_driver);



                //Performing the mouse hover action on the target element.
                action2.MoveToElement(policyOptionElement3).Perform();
                Delay(2);

                //Click on Reinstate
                _driver.FindElement(By.XPath("//*[@id='m0t0']/tbody/tr[10]/td/div/div[3]/a/img")).Click();
                Delay(2);



                
                 SelectElement oSelect2 = new SelectElement(_driver.FindElement(By.Name("frmReason")));

                oSelect2.SelectByValue("ReinstateReason2");
                Delay(2);


                //Click submit
                _driver.FindElement(By.Name("btnctcrereinstatecsu5")).Click();
                Delay(4);


                //Click submit
                _driver.FindElement(By.Name("btnctcrereinstatecsu2")).Click();
                Delay(5);



                //Contract Status validation 

                var StatusInForce = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr[2]/td[2]/u/font")).Text;

                
                
               Assert.IsTrue(StatusInForce.Equals("In-Force", StringComparison.CurrentCultureIgnoreCase));
                 
                Delay(3);

                if (StatusInForce == "In-Force")
                {
                    results = "Passed";
                }
                else
                {
                    results = "Failed";
                }


             


                base.writeResultsToExcell(results, sheet, "ReInstate");


            }

            catch (Exception ex)
   
            {

                DisconnectBrowser();

                throw ex;

            }

        }
        [Category("IncreaseSumAssured")]
        private void IncreaseSumAssured()
        {

            try

            {
                String test_url_3 = "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?";
                String test_url_4_title = "DateTime Picker";
                IJavaScriptExecutor js2 = (IJavaScriptExecutor)_driver;

                string contRef = base.GetPolicyNoFromExcell(sheet, "IncreaseSumAssured");

                string results = "";

                string date = DateTime.Today.ToString("g");


                policySearch(contRef);

                Delay(3);

                var contractPrem = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv9']/table/tbody/tr[2]/td[2]")).Text;



                //Click on user  component
                _driver.FindElement(By.Name("fccComponentDescription1")).Click();
                Delay(2);



                IWebElement policyOptionElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));


                //Creating object of an Actions class
                Actions action = new Actions(_driver);

                   

                //Performing the mouse hover action on the target element.
                action.MoveToElement(policyOptionElement).Perform();

                //Click on options
                _driver.FindElement(By.XPath("//*[@id='m0t0']/tbody/tr[1]/td/div/div[3]/a/img")).Click();

                Delay(5);

                SelectElement oSelect4 = new SelectElement(_driver.FindElement(By.Name("frmSPAmount")));

                oSelect4.SelectByValue("40000");
                Delay(2);

                //Click on calender
                _driver.FindElement(By.XPath("//*[@id='frmCbmcc']/tbody/tr[2]/td[2]/a/span/img")).Click();
                Delay(2);



                Assert.AreEqual(2, _driver.WindowHandles.Count);



                var newWindowHandle = _driver.WindowHandles[1];
                Assert.IsTrue(!string.IsNullOrEmpty(newWindowHandle));



                /* Assert.AreEqual(driver.SwitchTo().Window(newWindowHandle).Url, "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?"); */
                string expectedNewWindowTitle = test_url_4_title;
                Assert.AreEqual(_driver.SwitchTo().Window(newWindowHandle).Title, expectedNewWindowTitle);



                //Click on arrow
                _driver.FindElement(By.Id("aIncYear")).Click();
                Delay(2);




                //Click on Date
                _driver.FindElement(By.Name("fcCal1")).Click();
                Delay(2);




                /* Return to the window with handle = 0 */
                _driver.SwitchTo().Window(_driver.WindowHandles[0]);
                Delay(4);



                //Click on next
                _driver.FindElement(By.Name("btncbmcc13")).Click();
                Delay(2);


                //Click on next
                _driver.FindElement(By.Name("btncbmcc17")).Click();
                Delay(2);

                // Click on finish
                _driver.FindElement(By.Name("btncbmcc23")).Click();
                Delay(5);


                //Arrange
                var expectedamount = "40,000.00";

                //find the result
                string actualResult2 = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv5']/table/tbody/tr[2]/td[9]")).Text;
                Assert.IsTrue(expectedamount.Equals(actualResult2, StringComparison.CurrentCultureIgnoreCase));

                Delay(3);


                var newPrem = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr/td[2]")).Text;



                //// Click on contract summa
                //_driver.FindElement(By.Name("cb_User_cbmct")).Click();
                //Delay(5);


                //// Click on contract summa
                //_driver.FindElement(By.Name("cb_User_cbmct")).Click();
                //Delay(5);

                if (Convert.ToDecimal(newPrem) > Convert.ToDecimal(contractPrem) )
                {
                    results = "Passed";
                }
                else
                {
                    results = "Failed";
                }


                base.writeResultsToExcell(results, sheet, "IncreaseSumAssured");

            }

            catch (Exception ex)

            {

                DisconnectBrowser();

                throw ex;

            }

        }

        [Category("ChangeCollectionM")]
        private void ChangeCollectionM()
        {

            try

            {

                String test_url_1 = "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?";
                String test_url_2_title = "MIP - Sanlam ARL - Warpspeed Lookup Window";
                IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;


                string contRef = base.GetPolicyNoFromExcell(sheet, "ChangeCollectionM");

                string results = "";

                string date = DateTime.Today.ToString("g");


                policySearch(contRef);

                Delay(3);



                //click on policy payer  
                _driver.FindElement(By.Name("fcRoleEntityLink3")).Click();



                IWebElement policyOptionElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));


                //Creating object of an Actions class
                Actions action = new Actions(_driver);



                //Performing the mouse hover action on the target element.
                action.MoveToElement(policyOptionElement).Perform();


                //Click on options
                _driver.FindElement(By.XPath("//*[@id='m0t0']/tbody/tr[3]/td/div/div[3]/a/img")).Click();


                SelectElement oSelect4 = new SelectElement(_driver.FindElement(By.Name("fcCollectionMethod")));

                oSelect4.SelectByValue("108978.19");
                Delay(5);

                //Click on EMPLOYEE NUMBER
                _driver.FindElement(By.Name("fcEmployeeNumber")).Click();
                Delay(5);

                //Click on EMPLOYEE NUMBER
                _driver.FindElement(By.Name("fcEmployeeNumber")).SendKeys("23452345");
                Delay(5);


                //Search employee
                _driver.FindElement(By.Name("fcEmployerButton")).Click();


                Assert.AreEqual(2, _driver.WindowHandles.Count);

                var newWindowHandle = _driver.WindowHandles[1];
                Assert.IsTrue(!string.IsNullOrEmpty(newWindowHandle));

                /* Assert.AreEqual(driver.SwitchTo().Window(newWindowHandle).Url, "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?"); */
                string expectedNewWindowTitle = test_url_2_title;
                Assert.AreEqual(_driver.SwitchTo().Window(newWindowHandle).Title, expectedNewWindowTitle);



                //Search employee
                _driver.FindElement(By.XPath("//*[@id='lkpResultsTable']/tbody/tr[17]")).Click();


                /* Return to the window with handle = 0 */
                _driver.SwitchTo().Window(_driver.WindowHandles[0]);
                Delay(4);



                //Click on options
                _driver.FindElement(By.Id("GBLbl-1")).Click();
                Delay(5);

                //Click on contract summary
                _driver.FindElement(By.Name("PF_User_Menu")).Click();
                
                //Click on contract summary
                _driver.FindElement(By.Name("cb_User_cbmct")).Click();
                


                var expectedcollectionM = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv9']/table/tbody/tr[5]/td[2]")).Text;

               
               // Assert.IsTrue(expectedcollectionM.Equals("Stop Order", StringComparison.CurrentCultureIgnoreCase));

                Delay(3);

                if (expectedcollectionM == "DebiCheck")
                {
                    results = "Passed";
                }
                else
                {
                    results = "Failed";
                }


                base.writeResultsToExcell(results, sheet, "ChangeCollectionM");


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
