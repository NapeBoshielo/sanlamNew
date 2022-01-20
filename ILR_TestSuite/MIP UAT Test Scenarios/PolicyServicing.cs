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


            //clickOnMainMenu();
            //Delay(2);
            //AddRolePlayer();
            //Delay(2);
            //clickOnMainMenu();
            //Delay(4);
            //TerminateRolePlayer();
            //Delay(2);
            //clickOnMainMenu();
            //Delay(4);
            //IncreaseSumAssured();
            //Delay(2);
            //clickOnMainMenu();
            //Delay(4);
            //DecreaseSumAssured();
            //Delay(2);
            //clickOnMainMenu();
            //Delay(4);
            //AddRole_NextMonth();
            //Delay(2);
            //clickOnMainMenu();
            //Delay(4);
            //TerminateRoleNext_month();
            //Delay(2);
            //clickOnMainMenu();
            //Delay(4);
            //PostDatedDowngrade();
            //Delay(2);
            //clickOnMainMenu();
            //Delay(4);
            //PostDatedUpgrade();
            //Delay(2);
            //clickOnMainMenu();
            //Delay(4);
            //IncreaseSumAssuredAge();
            //Delay(2);
            //clickOnMainMenu();
            //Delay(4);
            //IncreaseSumAssured();
            //Delay(2);
            //clickOnMainMenu();
            //Delay(4);
            //DecreaseSumAssured();
            //Delay(2);
            //clickOnMainMenu();
            //Delay(4);
            //RemovalOfNonCompulsoryLife();
            //Delay(2);
           clickOnMainMenu();
            Delay(4);
            ChangeCollectionMeth();
            //Delay(2);
            //clickOnMainMenu();
            //Delay(2);
            //ReInstate();
            //Delay(4);
            //clickOnMainMenu();
            //Delay(2);
            //ChangeLifeAssured();
            //Delay(2);
            //clickOnMainMenu();
            //Delay(4);
            //AddaLife();
            //Delay(2);
            //clickOnMainMenu();
            //Delay(2);
            //addBeneficiary();
            //Delay(4);
            clickOnMainMenu();
            Delay(4);
            ChangeCollectionM();





        }
        private void clickOnMainMenu()
        {
            try
            {
                //find the contract search option
                var search = _driver.FindElement(By.Name("alf-ICF8_00000222"));
            }
            catch
            {
                //clickOnMainMenu
                _driver.FindElement(By.Name("CBWeb")).Click();
                
            }
        }

        private void PostDatedDowngrade()
        {
            try
            {
                string contRef = base.GetPolicyNoFromExcell(sheet, "PostDatedDowngrade");

                string results = "";

                string date = DateTime.Today.ToString("g");

                var currentSumAssured = "";



                policySearch(contRef);

                Delay(3);

                var contractPrem = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv9']/table/tbody/tr[2]/td[2]")).Text;



                //Click on user  component
                _driver.FindElement(By.Name("fccComponentDescription1")).Click();


                Delay(4);
                //Get The current Sum Assured for the life assured
                currentSumAssured = _driver.FindElement(By.XPath("//*[@id='frmCbmcc']/tbody/tr[8]/td[4]")).Text;


                Delay(2);
                IWebElement policyOptionElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));


                //Creating object of an Actions class
                Actions action = new Actions(_driver);



                //Performing the mouse hover action on the target element.
                action.MoveToElement(policyOptionElement).Perform();

                //Click on options
                _driver.FindElement(By.XPath("//*[@id='m0t0']/tbody/tr[1]/td/div/div[3]/a/img")).Click();



                Delay(3);
                var year = DateTime.Now.Year;
                var month = ((DateTime.Now.Month) + 1).ToString();

                if (month == "13")
                {
                    month = "01";

                }
                else if (Convert.ToInt32(month) < 10)
                {
                    month = "0" + month;
                }


                var day = "01";
                var dt = "" + year + month + day;
                _driver.FindElement(By.Name("frmCCStartDate")).Clear();
                Delay(3);
                _driver.FindElement(By.Name("frmCCStartDate")).SendKeys(dt);
                var newSumAssured = "";


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



                //Click on next
                _driver.FindElement(By.Name("btncbmcc13")).Click();
                Delay(2);


                //Click on next
                _driver.FindElement(By.Name("btncbmcc17")).Click();
                Delay(2);

                // Click on finish
                _driver.FindElement(By.Name("btncbmcc23")).Click();
                Delay(3);



                var expectedPrem = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv9']/table/tbody/tr[2]/td[4]")).Text;

                var eventDescription = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv9']/table/tbody/tr[2]/td[3]")).Text;



                if (Convert.ToDecimal(expectedPrem) < Convert.ToDecimal(contractPrem) && eventDescription == "Post Dated Downgrade")
                {
                    results = "Passed";
                }
                else
                {
                    results = "Failed";
                }


                base.writeResultsToExcell(results, sheet, "PostDatedDowngrade");


            }
            catch (Exception ex)
            {

                throw;
            }
        }

        private void PostDatedUpgrade()
        {
            try
            {
                string contRef = base.GetPolicyNoFromExcell(sheet, "PostDatedUpgrade");

                string results = "";

                string date = DateTime.Today.ToString("g");

                var currentSumAssured = "";



                policySearch(contRef);

                Delay(3);

                var contractPrem = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv9']/table/tbody/tr[2]/td[2]")).Text;



                //Click on user  component
                _driver.FindElement(By.Name("fccComponentDescription1")).Click();


                Delay(4);
                //Get The current Sum Assured for the life assured
                currentSumAssured = _driver.FindElement(By.XPath("//*[@id='frmCbmcc']/tbody/tr[8]/td[4]")).Text;


                Delay(2);
                IWebElement policyOptionElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));


                //Creating object of an Actions class
                Actions action = new Actions(_driver);



                //Performing the mouse hover action on the target element.
                action.MoveToElement(policyOptionElement).Perform();
                Delay(4);
                //Click on options
                _driver.FindElement(By.XPath("//*[@id='m0t0']/tbody/tr[1]/td/div/div[3]/a/img")).Click();



                Delay(3);
                var year = DateTime.Now.Year;
                var month = ((DateTime.Now.Month) + 1).ToString();

                if (month == "13")
                {
                    month = "01";

                }
                else if (Convert.ToInt32(month) < 10)
                {
                    month = "0" + month;
                }
                var day = "01";
                var dt = "" + year + month + day;
                _driver.FindElement(By.Name("frmCCStartDate")).Clear();
                Delay(3);
                _driver.FindElement(By.Name("frmCCStartDate")).SendKeys(dt);
                var newSumAssured = "";
                //Do a  upgrade on current sum assured by 5000
                if (Convert.ToInt32(currentSumAssured) > 10000 || Convert.ToInt32(currentSumAssured) == 10000)
                {
                    newSumAssured = (Convert.ToInt32(currentSumAssured) + 10000).ToString();
                }
                else
                {
                    newSumAssured = (10000).ToString();
                }

                SelectElement oSelect = new SelectElement(_driver.FindElement(By.Name("frmSPAmount")));

                oSelect.SelectByValue(newSumAssured);



                //Click on next
                _driver.FindElement(By.Name("btncbmcc13")).Click();
                Delay(2);


                //Click on next
                _driver.FindElement(By.Name("btncbmcc17")).Click();
                Delay(2);

                // Click on finish
                _driver.FindElement(By.Name("btncbmcc23")).Click();
                Delay(5);




                var expectedPrem = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv9']/table/tbody/tr[2]/td[4]")).Text;

                var eventDescription = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv9']/table/tbody/tr[2]/td[3]")).Text;



                if (Convert.ToDecimal(expectedPrem) > Convert.ToDecimal(contractPrem) && eventDescription == "Post Dated Upgrade")
                {
                    results = "Passed";
                }
                else
                {
                    results = "Failed";
                }



                base.writeResultsToExcell(results, sheet, "PostDatedUpgrade");


            }
            catch (Exception ex)
            {

                throw;
            }
        }
        private void AddRolePlayer()
        {

            try
            {
                IJavaScriptExecutor js2 = (IJavaScriptExecutor)_driver;

                string contRef = base.GetPolicyNoFromExcell(sheet, "AddRolePlayer");

                string results = "";
                string title = "", first_name = "", surname = "", initials = "", dob = "", gender = "", id_number = "", relationship = "";

                Delay(3);
                policySearch(contRef);

                Delay(3);
                //Redate
                redate();


                Delay(3);

                //click add role player

                _driver.FindElement(By.Name("btnAddRolePlayer")).Click();

                //Select role
                IWebElement selectRole = _driver.FindElement(By.Name("frmRoleObj"));
                SelectElement s = new SelectElement(selectRole);
                s.SelectByIndex(4);
                Delay(3);
                //Click calendar

                _driver.FindElement(By.XPath("//*[@id='frmCbmre']/tbody/tr[2]/td[2]/a")).Click();
                Delay(3);
                Assert.AreEqual(2, _driver.WindowHandles.Count);

                //switch to window
                var childwindow = _driver.WindowHandles[1];
                _driver.SwitchTo().Window(childwindow);



                // click calendar month
                Actions act = new Actions(_driver);
                IWebElement ele = _driver.FindElement(By.XPath("//*[@id='aIncYear']/img"));
                act.Click(ele).Perform();
                _driver.FindElement(By.XPath("//*[@id='aIncYear']")).Click();

                _driver.FindElement(By.XPath("//a[@name='fcCal1']")).Click();


                var parent = _driver.WindowHandles[0];
                _driver.SwitchTo().Window(parent);
                Delay(4);


                _driver.FindElement(By.XPath("//*[@id='GBLbl-4']/span/a")).Click();
                //assert
                string url = _driver.Url;
                Assert.AreEqual(url, "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?");

                //click next to enter new role player
                _driver.FindElement(By.XPath("//*[@id='GBLbl-5']/span/a")).Click();

                using (OleDbConnection con = new OleDbConnection(base._connString))
                {
                    try
                    {
                        con.Open();

                        String command = "SELECT * FROM [AddRolePlayer$]";

                        OleDbCommand cmd = new OleDbCommand(command, con);

                        OleDbDataAdapter adapt = new OleDbDataAdapter();
                        adapt.SelectCommand = cmd;

                        DataSet ds = new DataSet("policies");
                        adapt.Fill(ds);
                        foreach (var row in ds.Tables[0].DefaultView)
                        {

                            title = ((System.Data.DataRowView)row).Row.ItemArray[0].ToString();
                            first_name = ((System.Data.DataRowView)row).Row.ItemArray[1].ToString();
                            surname = ((System.Data.DataRowView)row).Row.ItemArray[2].ToString();
                            initials = ((System.Data.DataRowView)row).Row.ItemArray[3].ToString();
                            dob = ((System.Data.DataRowView)row).Row.ItemArray[4].ToString();
                            gender = ((System.Data.DataRowView)row).Row.ItemArray[5].ToString();
                            id_number = ((System.Data.DataRowView)row).Row.ItemArray[6].ToString();
                            relationship = ((System.Data.DataRowView)row).Row.ItemArray[7].ToString();
                            break;
                        }


                    }
                    catch (Exception ex)

                    {
                        throw ex;

                    }
                    con.Close();
                    con.Dispose();

                }

                //enter id from excel
                /*  _driver.FindElement(By.Name("frmIDNumber")).SendKeys(idno);
                  Delay(2);
                  _driver.FindElement(By.Name("frmPersonObjLkpImg")).Click();

                  Delay(4);

                  var child = _driver.WindowHandles[1];
                  _driver.SwitchTo().Window(child);
                  Delay(2);

                  _driver.FindElement(By.XPath(" //*[@id='lkpResultsTable']/tbody/tr[2]")).Click();

                  Delay(2);
                  var parentt = _driver.WindowHandles[0];
                  _driver.SwitchTo().Window(parentt);

                  Delay(2);
                //click next

                  _driver.FindElement(By.XPath("//*[@id='GBLbl-5']/span/a")).Click();
                  Delay(2);
                */

                //enter initials
                _driver.FindElement(By.Name("frmPersonInitials")).Clear();
                _driver.FindElement(By.Name("frmPersonInitials")).SendKeys(initials);
                Delay(2);
                //enter name
                _driver.FindElement(By.Name("frmPersonFirstName")).Clear();
                _driver.FindElement(By.Name("frmPersonFirstName")).SendKeys(first_name);
                Delay(2);
                //enter surname
                _driver.FindElement(By.Name("frmPersonLastName")).Clear();
                _driver.FindElement(By.Name("frmPersonLastName")).SendKeys(surname);
                Delay(2);
                //enter


                _driver.FindElement(By.Name("frmPersonIDNumber")).SendKeys(id_number);
                Delay(2);
                //enter
                _driver.FindElement(By.Name("frmPersonDateOfBirth")).SendKeys(dob);
                Delay(2);
                //marital status
                IWebElement merital = _driver.FindElement(By.Name("frmPersonMaritalStatus"));
                SelectElement iselect = new SelectElement(merital);
                iselect.SelectByIndex(1);


                //Select gender
                IList<IWebElement> rdos = _driver.FindElements(By.XPath("//input[@name='frmPersonGender']"));

                foreach (IWebElement radio in rdos)
                {

                    if (radio.GetAttribute("value").Equals("er_AcPerGenMal"))
                    {

                        radio.Click();
                        break;

                    }
                    else
                    {

                        radio.FindElement(By.XPath("//*[@id='frmSubCbmre']/tbody/tr[1]/td[4]/table/tbody/tr/td[3]/input")).Click();

                    }


                }

                //Select Relationship
                var V_relationship = "";
                switch (relationship)
                {

                    case "Additional parent":
                        V_relationship = "951372577.488";
                        break;
                    case "Spouse":
                        V_relationship = "854651144.248";
                        break;
                    case "Additional child":
                        V_relationship = "905324120.488";
                        break;
                    case "Child":
                        V_relationship = "905324138.488";
                        break;
                    case "Parent":
                        V_relationship = "347901097.188";
                        break;

                    case "Brother":
                        V_relationship = "951371842.488";
                        break;

                }

                IWebElement relation = _driver.FindElement(By.Name("frmRelationshipCodeObj"));
                SelectElement oselect = new SelectElement(relation);
                oselect.SelectByValue(V_relationship);
                //Title
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

                SelectElement oSelect = new SelectElement(_driver.FindElement(By.Name("frmPersonTitle")));
                oSelect.SelectByValue(value);




                string url2 = _driver.Url;
                Assert.AreEqual(url2, "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?");

                Delay(2);
                //save
                _driver.FindElement(By.XPath(" //*[@id='GBLbl-5']/span/a")).Click();
                Delay(2);
                //Roleplayer ID
                var LifeA_ID = _driver.FindElement(By.XPath(" //*[@id='frmSubCbmre']/tbody/tr[4]/td[4]")).Text;

                //validation
                if (id_number == LifeA_ID)
                {
                    results = "Passed";

                }
                else

                {
                    results = "Fail";

                }

                base.writeResultsToExcell(results, sheet, "AddRolePlayer");


                clickOnMainMenu();

                Delay(2);
                //click contract summary
                _driver.FindElement(By.XPath(" //*[@id='t0_754']/table/tbody/tr/td[3]/a")).Click();

                //*[@id="t0_754"]/table/tbody/tr/td[3]/a
                Delay(2);
                //click on add componet
                _driver.FindElement(By.XPath("   //*[@id='GBLbl-5']/span/a")).Click();


                //Delay(2);
                //  select component
                // IWebElement component = _driver.FindElement(By.Name("frmComponentObj"));
                //SelectElement selecComp = new SelectElement(component);
                //selecComp.SelectByIndex(1);

                Delay(2);

                IWebElement parentcomponent = _driver.FindElement(By.Name("frmParentComponentObj"));
                SelectElement selecCom = new SelectElement(parentcomponent);
                selecCom.SelectByIndex(1);
                Delay(2);

                _driver.FindElement(By.XPath("//*[@id='GBLbl-6']/span/a")).Click();


                Delay(2);

                _driver.FindElement(By.XPath("//*[@id='GB-6']")).Click();


                //Select date
                // click 
                // _driver.FindElement(By.XPath("//*[@id='frmCbmcc']/tbody/tr[2]/td[2]/a/span/img")).Click();


                //Delay(2);

                //var compchild = _driver.WindowHandles[1];
                //_driver.SwitchTo().Window(compchild);
                //  Delay(2);
                //_driver.FindElement(By.XPath("/html/body/form/table/tbody/tr[4]/td/table/tbody/tr[2]/td[4]/a")).Click();
                // Delay(2);
                // var cparent = _driver.WindowHandles[0];
                // _driver.SwitchTo().Window(cparent);

                // Delay(3);

                //select sumassured
                //string url3 = _driver.Url;
                //Assert.AreEqual(url3, "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?");



                //IWebElement sumassured = _driver.FindElement(By.Name("frmSPAmount"));
                //SelectElement sum = new SelectElement(sumassured);
                //sum.SelectByIndex(3);

                //Delay(3);

                //  _driver.FindElement(By.XPath("//*[@id='GBLbl-6']/span/a")).Click();
                // Delay(3);

                //click role player
                //Actions acts = new Actions(_driver);
                //IWebElement elem = _driver.FindElement(By.Name("frmRolePlayers"));
                //act.DoubleClick(elem).Perform();



                IWebElement elem = _driver.FindElement(By.XPath("//*[@id='frmCbmcc']/tbody/tr[6]/td[2]/select"));
                SelectElement roleP = new SelectElement(elem);
                roleP.SelectByIndex(0);
                Delay(3);


                //next
                _driver.FindElement(By.XPath(" //*[@id='GBLbl-7']/span/a")).Click();

                //Validate roleplayer ID number
                Delay(2);

                var RoleINno = _driver.FindElement(By.XPath("//*[@id='frmCbmcc']/tbody/tr[9]/td[2]")).Text;

                var Ridno = RoleINno.Split(" ")[RoleINno.Split(" ").Length - 1].ToString();
                var ID = Ridno.Substring(1, 13);

                Assert.AreEqual(id_number, ID);

                Delay(2);

                _driver.FindElement(By.XPath("//*[@id='GBLbl-7']/span/a")).Click();

                Delay(2);

            }
            catch (Exception) { throw; }




        }
        private void TerminateRolePlayer()
        {

            try
            {

                IJavaScriptExecutor js2 = (IJavaScriptExecutor)_driver;

                string contRef = base.GetPolicyNoFromExcell(sheet, "TerminateRoleP");

                string results = "";

                // _driver.FindElement(By.Name("CBWeb")).Click();
                // _driver.FindElement(By.XPath(" //*[@id='t0_754']/table/tbody/tr/td[3]/a")).Click();
                Delay(2);
                policySearch(contRef);

                for (int i = 2; i < 23; i++)
                {
                    var txt = _driver.FindElement(By.XPath($"//*[@id='CntContentsDiv11']/table/tbody/tr[{i.ToString()}]/td[1]/span")).Text;
                    var relationship = _driver.FindElement(By.XPath($"//*[@id='CntContentsDiv11']/table/tbody/tr[{i.ToString()}]/td[3]")).Text;

                    if (txt == "Life Assured" && relationship != "Self")
                    {
                        _driver.FindElement(By.XPath($"//*[@id='CntContentsDiv11']/table/tbody/tr[{i.ToString()}]/td[1]/span")).Click();
                        break;
                    }

                }
                Delay(3);

                // SS44244430


                String expected = "Are you sure you want to terminate this role";

                _driver.FindElement(By.Name("btnTerminate")).Click();

                String alerttext = _driver.SwitchTo().Alert().Text;
                _driver.SwitchTo().Alert().Accept();
                Assert.AreEqual(expected, alerttext);

                //change effective to date
                Delay(2);
                var effectivefrom = _driver.FindElement(By.Name("frmEffectiveFromDate"));
                var effectvalue = effectivefrom.GetAttribute("value");

                Delay(1);
                _driver.FindElement(By.Name("frmEffectiveToDate")).Clear();
                _driver.FindElement(By.Name("frmEffectiveToDate")).SendKeys(effectvalue);

                Delay(2);
                //click terminate

                _driver.FindElement(By.Name("btnTerminate")).Click();

                //validation
                if (_driver.Url == "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?")
                {
                    results = "Passed";

                }
                else

                {
                    results = "Fail";

                }

                base.writeResultsToExcell(results, sheet, "TerminateRoleP");





            }
            catch (Exception) { throw; }


        }

        private void redate()
        {

            string redate_date = "";
            //click on policy options
            var policyele = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));

            Actions act = new Actions(_driver);

            act.MoveToElement(policyele).Perform();

            Delay(1);
            //click on redate policy
            _driver.FindElement(By.XPath("//*[@id='m0t0']/tbody/tr[4]/td/div/div[3]/a/img")).Click();

            //get value of attribute 
            Delay(3);
            var element_value = _driver.FindElement(By.XPath("//*[@id='frmCommencementDate']"));
            string rdat = element_value.GetAttribute("value");

            using (OleDbConnection con = new OleDbConnection(base._connString))
            {
                try
                {
                    con.Open();

                    String command = "SELECT * FROM [AddRolePlayer$]";

                    OleDbCommand cmd = new OleDbCommand(command, con);

                    OleDbDataAdapter adapt = new OleDbDataAdapter();
                    adapt.SelectCommand = cmd;

                    DataSet ds = new DataSet("policies");
                    adapt.Fill(ds);
                    foreach (var row in ds.Tables[0].DefaultView)
                    {

                        redate_date = ((System.Data.DataRowView)row).Row.ItemArray[8].ToString();
                        //next date
                        if (redate_date == rdat)
                        {

                            _driver.FindElement(By.XPath("//*[@id='GBLbl-3']/span/a")).Click();

                        }
                        else
                        {
                            //current month
                            _driver.FindElement(By.XPath("//*[@id='frmctcremaint']/tbody/tr[7]/td[2]/a/span/img")).Click();



                            var cal = _driver.WindowHandles[1];
                            _driver.SwitchTo().Window(cal);

                            _driver.FindElement(By.XPath("//*[@id='aDecYear']/img")).Click();





                            _driver.FindElement(By.XPath("//a[@name='fcCal1']")).Click();

                            var cal2 = _driver.WindowHandles[0];
                            _driver.SwitchTo().Window(cal2);

                            Delay(2);
                            _driver.FindElement(By.Name("btnctcremik5")).Click();

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

        }

        private void AddRole_NextMonth()
        {

            try
            {
                IJavaScriptExecutor js2 = (IJavaScriptExecutor)_driver;

                string contRef = base.GetPolicyNoFromExcell(sheet, "AddRoleNext_month");
                string title = "", first_name = "", surname = "", initials = "", dob = "", gender = "", id_number = "", relationship = "";
                string results = "";
                string idno = "";
                policySearch(contRef);
                Delay(2);

                //click add role

                _driver.FindElement(By.Name("btnAddRolePlayer")).Click();


                //Select role
                IWebElement selectRole = _driver.FindElement(By.Name("frmRoleObj"));
                SelectElement s = new SelectElement(selectRole);
                s.SelectByIndex(4);
                Delay(3);
                //Click calendar

                //_driver.FindElement(By.XPath("//*[@id='frmCbmre']/tbody/tr[2]/td[2]/a")).Click();
                // Delay(3);
                // Assert.AreEqual(2, _driver.WindowHandles.Count);

                // var childwindow = _driver.WindowHandles[1];
                // _driver.SwitchTo().Window(childwindow);
                //Delay(2);
                //click
                //Actions act = new Actions(_driver);
                //IWebElement ele = _driver.FindElement(By.XPath("//*[@id='aIncYear']/img"));
                //act.DoubleClick(ele).Perform();
                //_driver.FindElement(By.XPath("//*[@id='aIncYear']")).Click();
                //_driver.FindElement(By.XPath("//a[@name='fcCal1']")).Click();
                //var parent = _driver.WindowHandles[0];
                //_driver.SwitchTo().Window(parent);
                //Delay(4);

                //click next
                _driver.FindElement(By.XPath("//*[@id='GBLbl-4']/span/a")).Click();

                string url = _driver.Url;
                Assert.AreEqual(url, "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?");

                //click next to enter new role player
                _driver.FindElement(By.XPath("//*[@id='GBLbl-5']/span/a")).Click();

                using (OleDbConnection con = new OleDbConnection(base._connString))
                {
                    try
                    {
                        con.Open();

                        String command = "SELECT * FROM [AddRoleNext_month$]";

                        OleDbCommand cmd = new OleDbCommand(command, con);

                        OleDbDataAdapter adapt = new OleDbDataAdapter();
                        adapt.SelectCommand = cmd;

                        DataSet ds = new DataSet("policies");
                        adapt.Fill(ds);
                        foreach (var row in ds.Tables[0].DefaultView)
                        {

                            title = ((System.Data.DataRowView)row).Row.ItemArray[0].ToString();
                            first_name = ((System.Data.DataRowView)row).Row.ItemArray[1].ToString();
                            surname = ((System.Data.DataRowView)row).Row.ItemArray[2].ToString();
                            initials = ((System.Data.DataRowView)row).Row.ItemArray[3].ToString();
                            dob = ((System.Data.DataRowView)row).Row.ItemArray[4].ToString();
                            gender = ((System.Data.DataRowView)row).Row.ItemArray[5].ToString();
                            id_number = ((System.Data.DataRowView)row).Row.ItemArray[6].ToString();
                            relationship = ((System.Data.DataRowView)row).Row.ItemArray[7].ToString();
                            break;
                        }


                    }
                    catch (Exception ex)

                    {
                        throw ex;

                    }
                    con.Close();
                    con.Dispose();

                }

                //enter id from excel
                /*  _driver.FindElement(By.Name("frmIDNumber")).SendKeys(idno);
                  Delay(2);
                  _driver.FindElement(By.Name("frmPersonObjLkpImg")).Click();

                  Delay(4);

                  var child = _driver.WindowHandles[1];
                  _driver.SwitchTo().Window(child);
                  Delay(2);

                  _driver.FindElement(By.XPath(" //*[@id='lkpResultsTable']/tbody/tr[2]")).Click();

                  Delay(2);
                  var parentt = _driver.WindowHandles[0];
                  _driver.SwitchTo().Window(parentt);

                  Delay(2);
                //click next

                  _driver.FindElement(By.XPath("//*[@id='GBLbl-5']/span/a")).Click();
                  Delay(2);
                */
                ///click next


                //enter initials
                _driver.FindElement(By.Name("frmPersonInitials")).Clear();
                _driver.FindElement(By.Name("frmPersonInitials")).SendKeys(initials);
                Delay(2);
                //enter name
                _driver.FindElement(By.Name("frmPersonFirstName")).Clear();
                _driver.FindElement(By.Name("frmPersonFirstName")).SendKeys(first_name);
                Delay(2);
                //enter surname
                _driver.FindElement(By.Name("frmPersonLastName")).Clear();
                _driver.FindElement(By.Name("frmPersonLastName")).SendKeys(surname);
                Delay(2);
                //enter


                _driver.FindElement(By.Name("frmPersonIDNumber")).SendKeys(id_number);
                Delay(2);
                //enter
                _driver.FindElement(By.Name("frmPersonDateOfBirth")).SendKeys(dob);
                Delay(2);
                //marital status
                IWebElement merital = _driver.FindElement(By.Name("frmPersonMaritalStatus"));
                SelectElement iselect = new SelectElement(merital);
                iselect.SelectByIndex(1);


                //Select gender
                IList<IWebElement> rdos = _driver.FindElements(By.XPath("//input[@name='frmPersonGender']"));

                foreach (IWebElement radio in rdos)
                {

                    if (radio.GetAttribute("value").Equals("er_AcPerGenMal"))
                    {

                        radio.Click();
                        break;

                    }
                    else
                    {

                        radio.FindElement(By.XPath("//*[@id='frmSubCbmre']/tbody/tr[1]/td[4]/table/tbody/tr/td[3]/input")).Click();

                    }


                }

                //Select Relationship
                var V_relationship = "";
                switch (relationship)
                {

                    case "Additional parent":
                        V_relationship = "951372577.488";
                        break;
                    case "Spouse":
                        V_relationship = "854651144.248";
                        break;
                    case "Additional child":
                        V_relationship = "905324120.488";
                        break;
                    case "Child":
                        V_relationship = "905324138.488";
                        break;
                    case "Parent":
                        V_relationship = "347901097.188";
                        break;

                    case "Brother":
                        V_relationship = "951371842.488";
                        break;

                }

                IWebElement relation = _driver.FindElement(By.Name("frmRelationshipCodeObj"));
                SelectElement oselect = new SelectElement(relation);
                oselect.SelectByValue(V_relationship);
                //Title
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

                SelectElement oSelect = new SelectElement(_driver.FindElement(By.Name("frmPersonTitle")));
                oSelect.SelectByValue(value);




                string url2 = _driver.Url;
                Assert.AreEqual(url2, "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?");





                Delay(2);
                //save
                _driver.FindElement(By.XPath(" //*[@id='GBLbl-5']/span/a")).Click();
                Delay(2);
                //Roleplayer ID
                var LifeA_ID = _driver.FindElement(By.XPath(" //*[@id='frmSubCbmre']/tbody/tr[4]/td[4]")).Text;

                //validation
                if (id_number == LifeA_ID)
                {
                    results = "Passed";

                }
                else

                {
                    results = "Fail";

                }

                base.writeResultsToExcell(results, sheet, "AddRoleNext_month");


                clickOnMainMenu();


                //click contract summary
                Delay(3);
                _driver.FindElement(By.XPath(" //*[@id='t0_754']/table/tbody/tr/td[3]/a")).Click();


                Delay(2);
                //click on componet
                _driver.FindElement(By.XPath("   //*[@id='GBLbl-5']/span/a")).Click();


                //Delay(2);
                ////  select component
                //IWebElement component = _driver.FindElement(By.Name("frmComponentObj"));
                //SelectElement selecComp = new SelectElement(component);
                //selecComp.SelectByIndex(1);

                Delay(2);

                IWebElement parentcomponent = _driver.FindElement(By.Name("frmParentComponentObj"));
                SelectElement selecCom = new SelectElement(parentcomponent);
                selecCom.SelectByIndex(1);
                Delay(2);

                _driver.FindElement(By.XPath("//*[@id='GBLbl-6']/span/a")).Click();


                Delay(2);
                //Select date
                // click 
                //_driver.FindElement(By.XPath("//*[@id='frmCbmcc']/tbody/tr[2]/td[2]/a/span/img")).Click();


                //Delay(2);

                //var compchild = _driver.WindowHandles[1];
                //_driver.SwitchTo().Window(compchild);
                //Delay(2);
                //_driver.FindElement(By.XPath("/html/body/form/table/tbody/tr[4]/td/table/tbody/tr[2]/td[4]/a")).Click();
                //Delay(2);
                //var cparent = _driver.WindowHandles[0];
                //_driver.SwitchTo().Window(cparent);

                //Delay(3);

                ////select sumassured
                //string url3 = _driver.Url;
                //Assert.AreEqual(url3, "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?");



                //IWebElement sumassured = _driver.FindElement(By.Name("frmSPAmount"));
                //SelectElement sum = new SelectElement(sumassured);
                //sum.SelectByIndex(3);

                //Delay(3);

                //next 
                _driver.FindElement(By.XPath("//*[@id='GBLbl-6']/span/a")).Click();
                Delay(3);

                //click role player
                IWebElement elem = _driver.FindElement(By.Name("frmRolePlayers"));
                SelectElement roleP = new SelectElement(elem);
                roleP.SelectByIndex(0);
                Delay(3);
                //next
                _driver.FindElement(By.XPath(" //*[@id='GBLbl-7']/span/a")).Click();

                //Validate roleplayer ID number


                var RoleINno = _driver.FindElement(By.XPath("//*[@id='frmCbmcc']/tbody/tr[9]/td[2]")).Text;

                var Ridno = RoleINno.Split(" ")[RoleINno.Split(" ").Length - 1].ToString();
                var ID = Ridno.Substring(1, 13);

                Assert.AreEqual(id_number, ID);

                _driver.FindElement(By.XPath("  //*[@id='GBLbl-7']/span/a")).Click();



            }
            catch (Exception) { throw; }


        }




        private void TerminateRoleNext_month()
        {

            try
            {

                IJavaScriptExecutor js2 = (IJavaScriptExecutor)_driver;

                string contRef = base.GetPolicyNoFromExcell(sheet, "TerminateRoleP_next");

                string results = "";

                // _driver.FindElement(By.Name("CBWeb")).Click();
                // _driver.FindElement(By.XPath(" //*[@id='t0_754']/table/tbody/tr/td[3]/a")).Click();
                Delay(2);
                policySearch(contRef);

                for (int i = 2; i < 23; i++)
                {
                    var txt = _driver.FindElement(By.XPath($"//*[@id='CntContentsDiv11']/table/tbody/tr[{i.ToString()}]/td[1]/span")).Text;
                    var relationship = _driver.FindElement(By.XPath($"//*[@id='CntContentsDiv11']/table/tbody/tr[{i.ToString()}]/td[3]")).Text;

                    if (txt == "Life Assured" && relationship != "Self")
                    {
                        _driver.FindElement(By.XPath($"//*[@id='CntContentsDiv11']/table/tbody/tr[{i.ToString()}]/td[1]/span")).Click();
                        break;
                    }

                }
                Delay(3);




                String expected = "Are you sure you want to terminate this role";

                _driver.FindElement(By.Name("btnTerminate")).Click();

                String alerttext = _driver.SwitchTo().Alert().Text;
                _driver.SwitchTo().Alert().Accept();
                Assert.AreEqual(expected, alerttext);

                //change effective to date
                Delay(2);
                var effectivefrom = _driver.FindElement(By.Name("frmEffectiveFromDate"));
                var effectvalue = effectivefrom.GetAttribute("value");

                Delay(1);
                _driver.FindElement(By.Name("frmEffectiveToDate")).Clear();
                _driver.FindElement(By.Name("frmEffectiveToDate")).SendKeys(effectvalue);

                Delay(2);
                //click terminate

                _driver.FindElement(By.Name("btnTerminate")).Click();

                //validation
                if (_driver.Url == "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?")
                {
                    results = "Passed";

                }
                else

                {
                    results = "Fail";

                }

                base.writeResultsToExcell(results, sheet, "TerminateRoleP_next");





            }
            catch (Exception) { throw; }


        }





        [Category("Add Beneficiary")]
        private void addBeneficiary()
        {
            var results = "Failed";
            string contRef = base.GetPolicyNoFromExcell(sheet, "AddBeneficiary");
            policySearch(contRef);
            Delay(2);
            var commDate = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr[6]/td[2]")).Text;
            _driver.FindElement(By.Name("btnAddRolePlayer")).Click();
            Delay(3);
            //Select role
            SelectElement oSelect = new SelectElement(_driver.FindElement(By.Name("frmRoleObj")));

            oSelect.SelectByValue("41667.19");
            Delay(2);
            //Effective date
            _driver.FindElement(By.Name("frmEffectiveFromDate")).Clear();
            Delay(2);
            _driver.FindElement(By.Name("frmEffectiveFromDate")).SendKeys(commDate);

            _driver.FindElement(By.Name("btncbmre7")).Click();
            Delay(3);
            _driver.FindElement(By.Name("frmIDNumber")).SendKeys("8604225772087");
            //8604225772087
            Delay(3);
            _driver.FindElement(By.Name("btncbmre13")).Click();
            Delay(10);

            oSelect = new SelectElement(_driver.FindElement(By.Name("frmRelationshipCodeObj")));

            oSelect.SelectByValue("905324144.488");
            Delay(3);
            _driver.FindElement(By.Name("btncbmre16")).Click();
            Delay(3);
            _driver.FindElement(By.Id("t0_749")).Click();
            Delay(3);
            clickOnMainMenu();
            Delay(3);

            _driver.FindElement(By.Name("2000175333.8")).Click();
            Delay(3);
            var row = 2;
            var maxRows = 23;
            for (int i = row; i < maxRows; i++)
            {

                string xpath = "//*[@id='CntContentsDiv11']/table/tbody/tr[" + i + "]/td[1]";
                try
                {
                    var role = _driver.FindElement(By.XPath(xpath)).Text;
                    if (role == "Beneficiary")
                    {
                        results = "Passed";
                        break;
                    }
                }
                catch (Exception)
                {
                    break;
                }


            }
            Delay(3);

            base.writeResultsToExcell(results, sheet, "AddBenefeciary");

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
                Delay(3);

                IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;


                Delay(4);

                //Get Current premium
                currentPremium = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv9']/table/tbody/tr[2]/td[2]")).Text;

                Delay(4);

                //Select the component
                _driver.FindElement(By.Name("fccComponentDescription1")).Click();




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

                var newSumAssured = "";
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

                var premuimfromRateTable = base.getPremuimFromRateTable(age, "ML", newSumAssured, "Safrican_Just_Funeral");

                Delay(4);
                _driver.FindElement(By.Name("btncbmcc23")).Click();
                Delay(6);

                //Get the new Premium
                //js.ExecuteScript("window.scrollBy(0,1000)", "");
                newPremium = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv5']/table/tbody/tr[2]/td[7]")).Text;

                //Do Age validation


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
        private void ChangeLifeAssured()
        {
            try
            {
                string contRef = base.GetPolicyNoFromExcell(sheet, "ChangeLifeAssured");

                string results = "";
                policySearch(contRef);

                string title = "", surname = "", MaritalStatus = "", EducationLevel = "", Department = "", Profession = "";

                //Extract data from excell
                using (OleDbConnection conn = new OleDbConnection(base._connString))
                {
                    try
                    {

                        // Open connection
                        conn.Open();
                        string cmdQuery = "SELECT * FROM [ChangeLifeData$]";

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
                            surname = ((System.Data.DataRowView)row).Row.ItemArray[1].ToString();
                            MaritalStatus = ((System.Data.DataRowView)row).Row.ItemArray[2].ToString();
                            EducationLevel = ((System.Data.DataRowView)row).Row.ItemArray[3].ToString();
                            Department = ((System.Data.DataRowView)row).Row.ItemArray[4].ToString();
                            Profession = ((System.Data.DataRowView)row).Row.ItemArray[5].ToString();

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
                //click on life Assured
                _driver.FindElement(By.Name("fcOwningEntityLink1")).Click();

                Delay(2);
                var oldMartialStatus = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv4']/table/tbody/tr[2]/td[4]")).Text;


                //click oN Change
                _driver.FindElement(By.Name("btnChangePerson")).Click();



                var value = "";
                switch (title)
                {
                    case ("Mr"):
                        value = "er_AcPerTitleMr";
                        break;
                    case ("Ms"):
                        value = "er_AcPerTitleMs";
                        break;
                    case ("Mrs"):
                        value = "er_AcPerTitleMrs";
                        break;
                    case ("Prf"):
                        value = "er_AcPerTitlePrf";
                        break;
                    case ("Dr"):
                        value = "er_AcPerTitleDr";
                        break;
                    case ("Adm"):
                        value = "er_AcPerTitleADM";
                        break;
                    case ("Adv"):
                        value = "er_AcPerTitleADV";
                        break;
                    case ("Capt"):
                        value = "er_AcPerTitleCAP";
                        break;
                    case ("Col"):
                        value = "er_AcPerTitleCOL";
                        break;
                    case ("Exec"):
                        value = "er_AcPerTitleEXE";
                        break;
                    case ("Gen"):
                        value = "er_AcPerTitleGEN";
                        break;
                    case ("Hon"):
                        value = "er_AcPerTitleHON";
                        break;
                    case ("Lt"):
                        value = "er_AcPerTitleLUI";
                        break;
                    case ("Maj"):
                        value = "er_AcPerTitleMAJ";
                        break;
                    case ("Messr"):
                        value = "er_AcPerTitleMES";
                        break;
                    case ("Pstr"):
                        value = "er_AcPerTitlePST";
                        break;
                    case ("Rev"):
                        value = "er_AcPerTitleREV";
                        break;
                    case ("Sir"):
                        value = "er_AcPerTitleSIR";
                        break;
                    case ("Brig"):
                        value = "er_AcPerTitleBRG";
                        break;
                    case ("Miss"):
                        value = "er_AcPerTitleMiss";
                        break;

                    default:
                        break;
                }

                SelectElement oSelect = new SelectElement(_driver.FindElement(By.Name("fcTitle")));
                //Select title
                oSelect.SelectByValue(value);
                Delay(2);

                //Insert Surname
                _driver.FindElement(By.Name("fcLastName")).Clear();
                _driver.FindElement(By.Name("fcLastName")).SendKeys(surname);

                Delay(2);

                //Insert MaritalStatus

                switch (MaritalStatus)
                {


                    case ("Single"):
                        value = "er_AcStaMarSin";
                        break;
                    case ("Married"):
                        value = "er_AcStaMarMar";
                        break;
                    case ("Divorced"):
                        value = "er_AcStaMarDiv";
                        break;
                    case ("Widowed"):
                        value = "er_AcStaMarWid";
                        break;
                    case ("Separated"):
                        value = "er_AcStaMarSep";
                        break;
                    case ("Unknown"):
                        value = "er_AcStaMarUnk";
                        break;
                    case ("Married (in Community of Property)"):
                        value = "er_AcStaMarMac";
                        break;
                    case ("Married (not in Community of Property)"):
                        value = "er_AcStaMarMnc";
                        break;
                    case ("Domestic Partnership/Co-habiting"):
                        value = "er_AcStaMarDpp";
                        break;

                    default:
                        break;

                }
                SelectElement oSelect2 = new SelectElement(_driver.FindElement(By.Name("fcMaritalStatus")));
                oSelect2.SelectByValue(value);
                Delay(2);


                //Insert EducationLevel
                switch (EducationLevel)
                {
                    case ("No Matric"):
                        value = "NMT";
                        break;
                    case ("Matric"):
                        value = "MTR";
                        break;
                    case ("Diploma (less than 3 years)"):
                        value = "DPL";
                        break;
                    case ("Diploma (3 years or more)"):
                        value = "DPM";
                        break;
                    case ("University/Undergraduate Degree"):
                        value = "UND";
                        break;
                    case ("Postgraduate Study"):
                        value = "PGD";
                        break;
                    default:
                        break; ;
                }

                SelectElement oSelect1 = new SelectElement(_driver.FindElement(By.Name("fcEducationLevel")));

                //Select title
                oSelect1.SelectByValue(value);
                Delay(2);

                //Department
                //Insert Surname
                _driver.FindElement(By.Name("fcDepartment")).Clear();
                _driver.FindElement(By.Name("fcDepartment")).SendKeys(Department);
                Delay(2);



                switch (Profession)
                {
                    case ("Accountant"):
                        value = "ACU";
                        break;
                    case ("Actor/Actress"):
                        value = "ACT";
                        break;
                    case ("Actuary"):
                        value = "ATY";
                        break;
                    case ("Adverting"):
                        value = "ADS";
                        break;
                    case ("Agriculture"):
                        value = "AGC";
                        break;
                    case ("Architect"):
                        value = "ARC";
                        break;
                    case ("Auditor"):
                        value = "AUD";
                        break;
                    case ("Banker"):
                        value = "BKR";
                        break;
                    case ("book keeper"):
                        value = "";
                        break;
                    case ("Broker"):
                        value = "BRK";
                        break;
                    case ("Doctor"):
                        value = "DTR";
                        break;
                    case ("Engineer"):
                        value = "EGR";
                        break;
                    case ("Human Resources"):
                        value = "HRC";
                        break;
                    case ("Import/Export"):
                        value = "IEX";
                        break;
                    case ("Information Technology"):
                        value = "ITC";
                        break;
                    case ("Insurance"):
                        value = "ISE";
                        break;
                    case ("Lawyer"):
                        value = "LAW";
                        break;
                    case ("Military"):
                        value = "MLY";
                        break;
                    case ("Pilot"):
                        value = "PLT";
                        break;

                    default:
                        break;
                }

                SelectElement oSelect3 = new SelectElement(_driver.FindElement(By.Name("fcProfession")));
                //Select title
                oSelect3.SelectByValue(value);
                Delay(2);

                //Click on the submit btn
                _driver.FindElement(By.Name("btnSubmit")).Click();


                Delay(2);
                var newMartialStatus = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv4']/table/tbody/tr[2]/td[4]")).Text;


                //Vaidation based Martial stat

                if ((oldMartialStatus) != (newMartialStatus))
                {
                    //Details sucessfully changed);
                    results = "Passed";
                }
                else
                {
                    results = "Failed";
                }


                base.writeResultsToExcell(results, sheet, "ChangeLifeAssured");

            }
            catch (Exception ex)
            {
                DisconnectBrowser();

                throw ex;
            }
        }

        [Category("Removal of Non-Compulsory life")]
        private void RemovalOfNonCompulsoryLife()
        {
            try
            {
                string contRef = base.GetPolicyNoFromExcell(sheet, "RemovaleOfNonCompulsoryLife");
                string results = "";
                var currentPremium = "";
                var newPremium = "";
                policySearch(contRef);
                Delay(2);
                //Get commencemet data
                var commDate = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr[6]/td[2]")).Text;
                Delay(2);
                //Get Current premium
                currentPremium = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv9']/table/tbody/tr[2]/td[2]")).Text;
                Delay(4);
                //Click on the component we want to terminate
                _driver.FindElement(By.Name("fccComponentDescription2")).Click();
                Delay(3);
                //Click on the terminate btn
                _driver.FindElement(By.Name("btncbmcc29")).Click();
                Delay(3);
                var commSplit = commDate.Split('/');
                commDate = "";
                foreach (var item in commSplit)
                {
                    commDate += item;
                }
                Delay(3);

                _driver.FindElement(By.Name("frmCCEndDate")).Clear();
                Delay(3);
                _driver.FindElement(By.Name("frmCCEndDate")).SendKeys(commDate);
                Delay(3);
                //Click on submit
                _driver.FindElement(By.Name("btncbmcc31")).Click();
                Delay(3);
                newPremium = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr/td[2]")).Text;
                Delay(3);
                results = Convert.ToDouble(newPremium) < Convert.ToDouble(currentPremium) ? "Passed" : "Failed";
                Delay(3);
                base.writeResultsToExcell(results, sheet, "RemovaleOfNonCompulsoryLife");
            }
            catch (Exception ex)
            {
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

                var newStatus = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr[2]/td[2]/u/font")).Text;

                if (newStatus == "Cancelled")
                {
                    //Assert.Pass("The policy was succesfully cancelled");
                    results = "Passed";
                }
                else
                {
                    results = "Failed";
                }



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

                CancelPolicy();
                //Contract Status validation 

                var Cancelled = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr[2]/td[2]/u/font")).Text;



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

                var currentSumAssured = "";
                var commDate = "";



                policySearch(contRef);

                Delay(3);

                //Get the Commencement date from contract summary screen
                commDate = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr[6]/td[2]")).Text;
                //Scroll Down
                Delay(2);

                IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;


                Delay(4);


                var contractPrem = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv9']/table/tbody/tr[2]/td[2]")).Text;



                //Click on user  component
                _driver.FindElement(By.Name("fccComponentDescription1")).Click();


                Delay(4);
                //Get The current Sum Assured for the life assured
                currentSumAssured = _driver.FindElement(By.XPath("//*[@id='frmCbmcc']/tbody/tr[8]/td[4]")).Text;


                Delay(2);
                IWebElement policyOptionElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));


                //Creating object of an Actions class
                Actions action = new Actions(_driver);



                //Performing the mouse hover action on the target element.
                action.MoveToElement(policyOptionElement).Perform();

                //Click on options
                _driver.FindElement(By.XPath("//*[@id='m0t0']/tbody/tr[1]/td/div/div[3]/a/img")).Click();

;

                //Date selection
                Delay(4);
                _driver.FindElement(By.Name("frmCCStartDate")).Clear();
                Delay(2);
                _driver.FindElement(By.Name("frmCCStartDate")).SendKeys(commDate);

                Delay(5);

                var newSumAssured = "";
                //Do a  upgrade on current sum assured by 5000
                if (Convert.ToInt32(currentSumAssured) > 10000 || Convert.ToInt32(currentSumAssured) == 10000)
                {
                    newSumAssured = (Convert.ToInt32(currentSumAssured) + 10000).ToString();
                }
                else
                {
                    newSumAssured = (60000).ToString();
                }

                SelectElement oSelect = new SelectElement(_driver.FindElement(By.Name("frmSPAmount")));

                oSelect.SelectByValue(newSumAssured);


                //Click on next
                _driver.FindElement(By.Name("btncbmcc13")).Click();
                Delay(2);


                //Click on next
                _driver.FindElement(By.Name("btncbmcc17")).Click();
                Delay(2);

                // Click on finish
                _driver.FindElement(By.Name("btncbmcc23")).Click();
                Delay(5);




                var newPrem = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr/td[2]")).Text;





                if (Convert.ToDecimal(newPrem) > Convert.ToDecimal(contractPrem))
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
        [Category("ChangeCollectionMeth")]
        private void ChangeCollectionMeth()
        {

            try

            {

                String test_url_1 = "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?";
                String test_url_2_title = "MIP - Sanlam ARL - Warpspeed Lookup Window";
                IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;


                string contRef = base.GetPolicyNoFromExcell(sheet, "ChangeCollectionM");

                string results = "";

                string date = DateTime.Today.ToString("g");

                string employee_number1 = "";

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

                using (OleDbConnection con = new OleDbConnection(base._connString))
                {
                    try
                    {
                        con.Open();

                        String command = "SELECT * FROM [CollectionM$]";

                        OleDbCommand cmd = new OleDbCommand(command, con);

                        OleDbDataAdapter adapt = new OleDbDataAdapter();
                        adapt.SelectCommand = cmd;

                        DataSet ds = new DataSet("policies");
                        adapt.Fill(ds);
                        foreach (var row in ds.Tables[0].DefaultView)
                        {


                            employee_number1 = ((System.Data.DataRowView)row).Row.ItemArray[1].ToString();

                            break;
                        }


                    }
                    catch (Exception ex)

                    {
                        throw ex;

                    }
                    con.Close();
                    con.Dispose();

                }

                //Click on EMPLOYEE NUMBER
                _driver.FindElement(By.Name("fcEmployeeNumber")).Click();
                Delay(5);

                //Click on EMPLOYEE NUMBER
                _driver.FindElement(By.Name("fcEmployeeNumber")).SendKeys(employee_number1);
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
                Delay(5);


                /* Return to the window with handle = 0 */
                _driver.SwitchTo().Window(_driver.WindowHandles[0]);
                Delay(5);


                //Click on submit
                _driver.FindElement(By.Id("GBLbl-1")).Click();
                Delay(5);





                var expectedcollectionM = _driver.FindElement(By.XPath("//*[@id='frmCbmre']/tbody/tr[8]/td[4]")).Text;




                Delay(3);

                if (expectedcollectionM == "Stop Order")
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

                string employee_number2 = "";

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

                using (OleDbConnection con = new OleDbConnection(base._connString))
                {
                    try
                    {
                        con.Open();

                        String command = "SELECT * FROM [CollectionM$]";

                        OleDbCommand cmd = new OleDbCommand(command, con);

                        OleDbDataAdapter adapt = new OleDbDataAdapter();
                        adapt.SelectCommand = cmd;

                        DataSet ds = new DataSet("policies");
                        adapt.Fill(ds);
                        foreach (var row in ds.Tables[0].DefaultView)
                        {

                         
                            employee_number2 = ((System.Data.DataRowView)row).Row.ItemArray[1].ToString();
                            
                            break;
                        }


                    }
                    catch (Exception ex)

                    {
                        throw ex;

                    }
                    con.Close();
                    con.Dispose();

                }

                //Click on EMPLOYEE NUMBER
                _driver.FindElement(By.Name("fcEmployeeNumber")).Click();
                Delay(5);

                //Click on EMPLOYEE NUMBER
                _driver.FindElement(By.Name("fcEmployeeNumber")).SendKeys(employee_number2);
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
                Delay(5);


                /* Return to the window with handle = 0 */
                _driver.SwitchTo().Window(_driver.WindowHandles[0]);
                Delay(5);


                //Click on submit
                _driver.FindElement(By.Id("GBLbl-1")).Click();
                Delay(5);





                var expectedcollectionM = _driver.FindElement(By.XPath("//*[@id='frmCbmre']/tbody/tr[8]/td[4]")).Text;


                

                Delay(3);

                if (expectedcollectionM == "Debi-Check")
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

        [Category("AddaLife")]
        private void AddaLife()
        {

            try

            {
                IJavaScriptExecutor js2 = (IJavaScriptExecutor)_driver;

                string contRef = base.GetPolicyNoFromExcell(sheet, "AddaLife");

                string results = "";
                string title = "", first_name = "", surname = "", initials = "", dob = "", gender = "", id_number = "", relationship = "";

                policySearch(contRef);


                var oldPrem = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv9']/table/tbody/tr[2]/td[2]")).Text;


                //click add policy

                _driver.FindElement(By.XPath("//*[@id='GBLbl-1']/span/a")).Click();

                //Select role
                IWebElement selectRole = _driver.FindElement(By.Name("frmRoleObj"));
                SelectElement s = new SelectElement(selectRole);
                s.SelectByIndex(4);
                // Delay(3);

                //Click calendar

                //_driver.FindElement(By.XPath("//*[@id='frmCbmre']/tbody/tr[2]/td[2]/a")).Click();
                //Delay(3);
                //Assert.AreEqual(2, _driver.WindowHandles.Count);

                //switch to window
                //var childwindow = _driver.WindowHandles[1];
                //_driver.SwitchTo().Window(childwindow);

                //Delay(2);

                ////clickback
                //Actions act = new Actions(_driver);
                //IWebElement ele = _driver.FindElement(By.XPath("//*[@id='aIncYear']/img"));
                //act.MoveToElement(ele).Click().Build().Perform();

                //_driver.FindElement(By.XPath("//a[@name='fcCal1']")).Click();


                //var parent = _driver.WindowHandles[0];
                //_driver.SwitchTo().Window(parent);
                Delay(4);


                _driver.FindElement(By.XPath("//*[@id='GBLbl-4']/span/a")).Click();
                //assert
                string url = _driver.Url;
                Assert.AreEqual(url, "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?");

                //click next to enter new role player
                _driver.FindElement(By.XPath("//*[@id='GBLbl-5']/span/a")).Click();

                using (OleDbConnection con = new OleDbConnection(base._connString))
                {
                    try
                    {
                        con.Open();

                        String command = "SELECT * FROM [AddaLife$]";

                        OleDbCommand cmd = new OleDbCommand(command, con);

                        OleDbDataAdapter adapt = new OleDbDataAdapter();
                        adapt.SelectCommand = cmd;

                        DataSet ds = new DataSet("policies");
                        adapt.Fill(ds);
                        foreach (var row in ds.Tables[0].DefaultView)
                        {

                            title = ((System.Data.DataRowView)row).Row.ItemArray[0].ToString();
                            first_name = ((System.Data.DataRowView)row).Row.ItemArray[1].ToString();
                            surname = ((System.Data.DataRowView)row).Row.ItemArray[2].ToString();
                            initials = ((System.Data.DataRowView)row).Row.ItemArray[3].ToString();
                            dob = ((System.Data.DataRowView)row).Row.ItemArray[4].ToString();
                            gender = ((System.Data.DataRowView)row).Row.ItemArray[5].ToString();
                            id_number = ((System.Data.DataRowView)row).Row.ItemArray[6].ToString();
                            relationship = ((System.Data.DataRowView)row).Row.ItemArray[7].ToString();




                            break;
                        }


                    }
                    catch (Exception ex)

                    {
                        throw ex;

                    }
                    con.Close();
                    con.Dispose();

                }



                //enter initials
                _driver.FindElement(By.Name("frmPersonInitials")).Clear();
                _driver.FindElement(By.Name("frmPersonInitials")).SendKeys(initials);
                Delay(2);
                //enter name
                _driver.FindElement(By.Name("frmPersonFirstName")).Clear();
                _driver.FindElement(By.Name("frmPersonFirstName")).SendKeys(first_name);
                Delay(2);
                //enter surname
                _driver.FindElement(By.Name("frmPersonLastName")).Clear();
                _driver.FindElement(By.Name("frmPersonLastName")).SendKeys(surname);
                Delay(2);
                //enter


                _driver.FindElement(By.Name("frmPersonIDNumber")).SendKeys(id_number);
                Delay(2);



                //enter
                _driver.FindElement(By.Name("frmPersonDateOfBirth")).Clear();
                _driver.FindElement(By.Name("frmPersonDateOfBirth")).SendKeys(dob);
                Delay(2);


                //marital status
                IWebElement merital = _driver.FindElement(By.Name("frmPersonMaritalStatus"));
                SelectElement iselect = new SelectElement(merital);
                iselect.SelectByIndex(1);


                //Select gender
                IList<IWebElement> rdos = _driver.FindElements(By.XPath("//input[@name='frmPersonGender']"));

                foreach (IWebElement radio in rdos)
                {

                    if (radio.GetAttribute("value").Equals("er_AcPerGenMal"))
                    {

                        radio.Click();
                        break;

                    }
                    else
                    {

                        radio.FindElement(By.XPath("//*[@id='frmSubCbmre']/tbody/tr[1]/td[4]/table/tbody/tr/td[3]/input")).Click();

                    }


                }

                //Select Relationship
                var V_relationship = "";
                switch (relationship)
                {

                    case "Additional parent":
                        V_relationship = "951372577.488";
                        break;
                    case "Spouse":
                        V_relationship = "854651144.248";
                        break;
                    case "Additional child":
                        V_relationship = "905324120.488";
                        break;
                    case "Child":
                        V_relationship = "905324138.488";
                        break;
                    case "Parent":
                        V_relationship = "347901097.188";
                        break;

                    case "Brother":
                        V_relationship = "951371842.488";
                        break;

                }

                IWebElement relation = _driver.FindElement(By.Name("frmRelationshipCodeObj"));
                SelectElement oselect = new SelectElement(relation);
                oselect.SelectByValue(V_relationship);
                //Title
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

                SelectElement oSelect = new SelectElement(_driver.FindElement(By.Name("frmPersonTitle")));
                oSelect.SelectByValue(value);

                string url2 = _driver.Url;
                Assert.AreEqual(url2, "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?");



                Delay(2);
                //save
                _driver.FindElement(By.XPath(" //*[@id='GBLbl-5']/span/a")).Click();
                Delay(2);



                clickOnMainMenu();



                //click contract summary
                _driver.FindElement(By.Name("2000175333.8")).Click();


                Delay(2);
                //click on componet
                _driver.FindElement(By.XPath("//*[@id='GBLbl-5']/span/a")).Click();


                Delay(2);

                IWebElement parentcomponent = _driver.FindElement(By.Name("frmParentComponentObj"));
                SelectElement selecCom = new SelectElement(parentcomponent);
                selecCom.SelectByIndex(1);
                Delay(2);

                _driver.FindElement(By.XPath("//*[@id='GBLbl-6']/span/a")).Click();






                SelectElement oSelect4 = new SelectElement(_driver.FindElement(By.Name("frmSPAmount")));

                oSelect4.SelectByValue("7500");
                Delay(2);

                //Click next
                _driver.FindElement(By.Name("btncbmcc2")).Click();
                Delay(2);


                //Click on next
                _driver.FindElement(By.Name("btncbmcc5")).Click();
                Delay(2);

                //Click on finish
                _driver.FindElement(By.Name("btncbmcc11")).Click();
                Delay(2);






                var newPrem = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr/td[2]")).Text;



                IJavaScriptExecutor js3 = (IJavaScriptExecutor)_driver;
                js3.ExecuteScript("window.scrollTo(0, 250)");


                if (Convert.ToDecimal(newPrem) > Convert.ToDecimal(oldPrem))
                {
                    results = "Passed";
                }
                else
                {
                    results = "Failed";
                }



                base.writeResultsToExcell(results, sheet, "AddaLife");

            }

            catch (Exception ex)

            {

                DisconnectBrowser();

                throw ex;

            }

        }

        private void IncreaseSumAssuredAge()
        {
            try
            {
                string contRef = base.GetPolicyNoFromExcell(sheet, "IncreaseSumAssuredAge");

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
                _driver.FindElement(By.Name("fccComponentDescription1")).Click();




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

                var newSumAssured = "";
                //Do a  upgrade on current sum assured by 5000
                if (Convert.ToInt32(currentSumAssured) > 10000 || Convert.ToInt32(currentSumAssured) == 10000)
                {
                    newSumAssured = (Convert.ToInt32(currentSumAssured) + 10000).ToString();
                }
                else
                {
                    newSumAssured = (60000).ToString();
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

                var premuimfromRateTable = base.getPremuimFromRateTable(age, "ML", newSumAssured, "Safrican_Just_Funeral");

                Delay(4);
                _driver.FindElement(By.Name("btncbmcc23")).Click();
                Delay(6);

                //Get the new Premium
                //js.ExecuteScript("window.scrollBy(0,1000)", "");
                newPremium = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv5']/table/tbody/tr[2]/td[7]")).Text;

                //Do Age validation
                string dateInput = "9905208629080";
                var splitted = dateInput.Substring(0, 2);


                //var rateTablePremium = base.getPremuimFromRateTable();
                Delay(2);
                clickOnMainMenu();
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
                base.writeResultsToExcell(results, sheet, "IncreaseSumAssuredAge");


            }
            catch (Exception)
            {

                throw;
            }
        }



        [Category("Policy Search")]
        public void policySearch(string contractRef = "")
        {

            IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;


            //Click on contract search 
            _driver.FindElement(By.Name("alf-ICF8_00000222")).Click();
            Delay(2);




            //Type in contract ref 

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
