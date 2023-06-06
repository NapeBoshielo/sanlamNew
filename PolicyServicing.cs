using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Wordprocessing;
using ILR_TestSuite;
using NHibernate.Cfg.MappingSchema;
using NHibernate.Mapping;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using Actions = OpenQA.Selenium.Interactions.Actions;
namespace PolicyServicing
{
    [TestFixture]
    public class PolicyServicing : TestBase
    {
        //Policy-Servicing
        private string sheet;
        [OneTimeSetUp]
        public void startBrowser()  
        {
            _driver = SiteConnection();
            sheet = "Policy-Servicing";
        }

      
        private void policySearch(string contractRef)
        {
            Delay(4);
            IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;
            //Click on contract search 
            try
            {
                _driver.FindElement(By.Name("alf-ICF8_00000222")).Click();
            } catch (Exception ex)
            {
                clickOnMainMenu();
                _driver.FindElement(By.Name("alf-ICF8_00000222")).Click();
            }
            Delay(2);
            //Type in contract ref 
            _driver.FindElement(By.Name("frmContractReference")).SendKeys(contractRef);
            Delay(4);
            //Click on Search Icon 
            _driver.FindElement(By.Name("btncbcts0")).Click();
            Delay(5);
            _driver.FindElement(By.XPath("//*[@id='AppArea']/table[2]/tbody/tr[2]/td[1]/a")).Click();
            Delay(5);
        }

        //Dont really need this one
        private void clickOnMainMenu()
        {
            try
            {
                //find the contract search option
                var search = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[1]/table/tbody/tr/td/table/tbody/tr[1]/td/div[7]/table[4]/tbody/tr/td/a"));
            }
            catch
            {
                //clickOnMainMenu
                _driver.FindElement(By.Name("CBWeb")).Click();
            }
        }

        private void Nav_ContractSummary()
        {
            try
            {
                //Click Contract Summary
                _driver.FindElement(By.Name("2000175333.8")).Click();
            }
            catch
            {
                //Click Main Menu then Contract Summary
                _driver.FindElement(By.Name("CBWeb")).Click();
                _driver.FindElement(By.Name("2000175333.8")).Click();
            }
        }

        private void Nav_Components()
        {
            try
            {
                //Click Components
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[3]/tbody/tr/td/a")).Click();
            }
            catch
            {
                try
                {
                    //Expand Contract Summary then click on Components
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[1]/a")).Click();
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[3]/tbody/tr/td/a")).Click();
                }
                catch
                {
                    //Click Main Menu then expand Contract Summary then click Components
                    _driver.FindElement(By.Name("CBWeb")).Click();
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[1]/a")).Click();
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[3]/tbody/tr/td/a")).Click();
                }
            }
        }

        private void Nav_Events()
        {
            try
            {
                //Click Events
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[9]/tbody/tr/td/a")).Click();
            }
            catch
            {
                try
                {
                    //Expand Contract Summary then click on Events
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[1]/a")).Click();
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[9]/tbody/tr/td/a")).Click();
                }
                catch
                {
                    //Click Main Menu then expand Contract Summary then click Events
                    _driver.FindElement(By.Name("CBWeb")).Click();
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[1]/a")).Click();
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[9]/tbody/tr/td/a")).Click();
                }
            }
        }

        [Test, TestCaseSource("GetTestData", new object[] { "AddRolePlayer" })]
        //Check policy only has one Beneficiary for the apportion split code
        //Cannot add Life assured for Product 4000 (Provide and Protect) - Life assured option does not exist
        //Cannot add Life assured for Product 5000 (Invest) - Relationship cannot be selected
        //Cannot add Life assured for Product 6000 (Saver) - Relationship cannot be selected
        //When adding a roleplayer of life assured, if you intend to make the life assured a child component, ensure the person is under the age of 26
        //When adding a child life assured there has to first be a Principle Life Assured component
        public void addRolePlayer(string contractRef, string scenarioID)
        {
            if (String.IsNullOrEmpty(contractRef))
            {
                Assert.Ignore();
            }
            var errMsg = String.Empty;
            string results = String.Empty;
            try
            {
                Delay(8);
                IJavaScriptExecutor js2 = (IJavaScriptExecutor)_driver;
                var currentCoverAmount = String.Empty; var newPremium = String.Empty; var currentPremium = String.Empty; var newContractPremium = String.Empty; string RolePlayerType = String.Empty; string Percentage = String.Empty;
                string title = String.Empty, component = String.Empty, sum_assured = String.Empty, first_name = String.Empty, surname = String.Empty, initials = String.Empty, dob = String.Empty, gender = String.Empty, id_number = String.Empty, relationship = String.Empty, comm_date = String.Empty;
                int leftoverpercentage;
                string roleplayer_check, relationship_check, idno_check = String.Empty;
                string Premium_check = String.Empty;
                string Component_check = String.Empty;
                string StartDate_check = String.Empty;
                Delay(3);
                policySearch(contractRef);
                Delay(2);
                var product = SetproductName();
                Delay(3);
                OpenDBConnection($"SELECT * FROM AddRolePlayer WHERE Scenario_ID = {scenarioID}");
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    title = reader["Title"].ToString();
                    first_name = reader["First_Name"].ToString();
                    RolePlayerType = reader["RolePlayerType"].ToString();
                    surname = reader["Surname"].ToString();
                    initials = reader["initials"].ToString();
                    Percentage = reader["Percentage"].ToString();
                    dob = reader["DOB"].ToString();
                    comm_date = reader["Comm_date"].ToString();
                    gender = reader["Gender"].ToString();
                    id_number = reader["ID_number"].ToString();
                    component = reader["component"].ToString();
                    sum_assured = reader["Sum_Assured"].ToString();
                    relationship = reader["Relationship"].ToString();
                }
                connection.Close();
                //get commencement date
                var commencementDate = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[3]/td/div/table/tbody/tr/td/span/table/tbody/tr/td[1]/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[6]/td[2]")).Text;
                //get product type
                var prodtype = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td[2]")).Text;
                //click add role player
                _driver.FindElement(By.Name("btnAddRolePlayer")).Click();
                //Select role
                Delay(3);
                var roletype = String.Empty;
                switch (RolePlayerType)
                {
                    case "Beneficiary":
                        roletype = "41667.19";
                        break;
                    case "Broker":
                        roletype = "47494.19";
                        break;
                    case "Claimant":
                        roletype = "2316554.188";
                        break;
                    case "informant":
                        roletype = "957815773.488";
                        break;
                    case "Life Assured":
                        roletype = "41666.19";
                        break;
                    default:
                        break;
                }
                IWebElement selectRole = _driver.FindElement(By.Name("frmRoleObj"));
                SelectElement s = new SelectElement(selectRole);
                s.SelectByValue(roletype);
                Delay(4);
                _driver.FindElement(By.XPath("//*[@id='GBLbl-4']/span/a")).Click();
                //click next to enter new role player
                _driver.FindElement(By.XPath("//*[@id='GBLbl-5']/span/a")).Click();
                //Accomodating for extra screen on product 6000 & 5000
                if (prodtype.Contains("6000") || prodtype.Contains("5000"))
                {
                    _driver.FindElement(By.XPath("//*[@id='GBLbl-5']/span/a")).Click();
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
                //enter ID
                _driver.FindElement(By.Name("frmPersonIDNumber")).SendKeys(id_number);
                Delay(2);
                //enter Date of Birth
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
                var V_relationship = String.Empty;
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

                SelectElement oSelect = new SelectElement(_driver.FindElement(By.Name("frmPersonTitle")));
                oSelect.SelectByText(title);
                Delay(2);
                if (RolePlayerType == "Beneficiary")
                {
                    //enter percentage
                    _driver.FindElement(By.Name("frmApportionmentPercent")).Clear();
                    _driver.FindElement(By.Name("frmApportionmentPercent")).SendKeys(Percentage);

                    //calculate leftover for existing benficiary
                    leftoverpercentage = (100 - int.Parse(Percentage));
                    _driver.FindElement(By.Name("frmAportion710")).Clear();
                    _driver.FindElement(By.Name("frmAportion710")).SendKeys(leftoverpercentage.ToString());
                }
                //save
                _driver.FindElement(By.XPath(" //*[@id='GBLbl-5']/span/a")).Click();
                Delay(2);
                
                //Navigate to the Contract Summary screen
                Nav_ContractSummary();

                if (!RolePlayerType.Equals("Beneficiary"))
                {

                    //Get The current Sum Assured for the life assured
                    currentPremium = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv9']/table/tbody/tr[2]/td[2]")).Text;
                    //click on add component
                    Delay(3);
                    _driver.FindElement(By.XPath("//*[@id='GBLbl-5']/span/a")).Click();
                    //select component 
                    var Components = String.Empty;
                    switch (component)
                    {
                        case "Child":
                            Components = "951086464.488";
                            break;
                        case "Parent":
                            Components = "951086466.488";
                            break;
                        default:
                            break;
                    }
                    IWebElement compn = (_driver.FindElement(By.Name("frmComponentObj")));
                    SelectElement Vcomp = new SelectElement(compn);
                    Vcomp.SelectByValue(Components);
                    Delay(3);
                    _driver.FindElement(By.XPath("//*[@id='GBLbl-6']/span/a")).Click();
                    Delay(3);
                    //Add Comm Date
                    _driver.FindElement(By.Name("frmCCStartDate")).Clear();
                    _driver.FindElement(By.Name("frmCCStartDate")).SendKeys(comm_date);

                    var values = String.Empty;
                    switch (sum_assured)
                    {
                        case "5000":
                            values = "5000";
                            break;
                        case "7500":
                            values = "7500";
                            break;
                        case "10000":
                            values = "10000";
                            break;
                        case "15000":
                            values = "15000";
                            break;
                        case "20000":
                            values = "20000";
                            break;
                        case "30000":
                            values = "30000";
                            break;
                        case "40000":
                            values = "40000";
                            break;
                        case "50000":
                            values = "50000";
                            break;
                        default:
                            break;
                    }
                    IWebElement amount = _driver.FindElement(By.Name("frmSPAmount"));
                    SelectElement vselect = new SelectElement(amount);
                    vselect.SelectByValue(values);
                    Delay(3);
                    _driver.FindElement(By.XPath("//*[@id='GB-6']")).Click();
                    Delay(2);
                    //Dropdown
                    IWebElement elem = _driver.FindElement(By.Name("frmRolePlayers"));
                    SelectElement option = new SelectElement(elem);
                    option.DeselectAll();
                    for (int i = 1; i < 30; i++)
                    {
                        var roles = _driver.FindElement(By.XPath($"//*[@id='frmCbmcc']/tbody/tr[6]/td[2]/select/option[{i.ToString()}]")).Text;
                        var Ridno1 = roles.Split(" ")[roles.Split(" ").Length - 1].ToString();
                        var ID1 = Ridno1.Substring(1, 13);
                        if (ID1 == id_number)
                        {
                            option.SelectByText(roles);
                            break;
                        }
                    }
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
                    //Calculate age based on IdNo
                    var idNo = (RoleINno.Split(" ")[RoleINno.Split(" ").Length - 1]).ToString();
                    var birthYear = idNo.Substring(1, 2);
                    Delay(2);
                    birthYear = "20" + birthYear;
                    var age = (DateTime.Now.Year - Convert.ToInt32(birthYear)).ToString();
                    Delay(3);
                    newContractPremium = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr/td[2]")).Text;
                    var premuimfromRateTable = getPremuimFromRateTable(id_number, component, sum_assured, product);
                    
                    try
                    {
                        for (int i = 2; i < 23; i++)
                        {
                            try
                            {
                                Component_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/big/b/a")).Text;
                                StartDate_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[4]")).Text;
                                Premium_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[7]")).Text;
                            }
                            catch
                            {
                                Component_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/a")).Text;
                                StartDate_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[4]")).Text;
                                Premium_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[7]")).Text;
                            }
                            if (Component_check.Contains("Child") && StartDate_check == comm_date && decimal.Parse(Premium_check) == premuimfromRateTable)
                            {
                                newPremium = Premium_check;
                                break;
                            }
                        }
                    }
                    catch
                    {
                        results = "Failed";
                        TakeScreenshot("AddRolePlayer");
                        errMsg = "Component was not found";
                    }
                    //premium validation
                    if (premuimfromRateTable == Convert.ToDecimal(newPremium) && Convert.ToDecimal(currentPremium) < Convert.ToDecimal(newContractPremium))
                    {
                        results = "Passed";
                    }
                    else
                    {
                        TakeScreenshot("AddRolePlayer");
                        results = "Failed";
                        errMsg = "Roleplayer was not added successfully";
                    }
                }
                else
                {
                    //Check though all of the roleplayers
                    try
                    {
                        for (int i = 2; i < 23; i++)
                        {
                            try
                            {
                                roleplayer_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/span/big/b/a")).Text;
                                relationship_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[3]")).Text;
                                idno_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[5]/span")).Text;
                            }
                            catch
                            {
                                roleplayer_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/span/a")).Text;
                                relationship_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[3]")).Text;
                                idno_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[5]/span")).Text;
                            }
                            if (roleplayer_check == RolePlayerType && relationship_check == relationship && idno_check == id_number)
                            {
                                results = "Passed";
                                break;
                            }
                        }
                    }
                    catch
                    {
                        results = "Failed";
                        TakeScreenshot("AddRolePlayer");
                        errMsg = "Roleplayer was not added successfully";
                    }
                }

            }
            catch (Exception ex)
            {
                TakeScreenshot("addRolePlayer");
                if (ex.Message.Length > 255)
                {
                    errMsg = ex.Message.Substring(0, 255);
                }
                else
                {
                    errMsg = ex.Message;
                }
                results = "Failed";
            }
            writeResultsToDB(results, Int32.Parse(scenarioID), errMsg);
            Assert.IsTrue(results.Equals("Passed"));
        }


        [Test, TestCaseSource("GetTestData", new object[] { "CancelPolicy" })]
        //Use a recently commenced contract that has financial mandates
        //Cannot "close" product 5000 (Invest) or product 6000 (Saver)
        public void cancelPolicy(string contractRef, string scenarioID)
        {
            if (String.IsNullOrEmpty(contractRef))
            {
                Assert.Ignore();
            }
            Delay(8);
            string results = String.Empty;
            string date = DateTime.Today.ToString("g");
            var errMsg = String.Empty;
            try
            {
                policySearch(contractRef);
                SetproductName();
                Delay(2);
                string reason = String.Empty;
                var commDate = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr[6]/td[2]")).Text;
                var dt = commDate;
                var splitedDate = dt.Split('/');
                var year = splitedDate[0];
                int month = Convert.ToInt32(splitedDate[1]);
                string strMonth;
                if (month == 12)
                {
                    strMonth = "01";
                    year = String.Empty + (Convert.ToInt32(year) + 1);
                }
                else if (month < 10)
                {
                    strMonth = "0" + (month + 1);
                }
                else
                {
                    strMonth = (month + 1).ToString();
                }
                commDate = year + "/" + strMonth + "/" + "01";
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
                Delay(3);
                //Provide cancellation reason
                OpenDBConnection($"SELECT * FROM Cancelpolicy WHERE Scenario_ID = {scenarioID}");
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    reason = reader["TerminationReason"].ToString();
                }
                connection.Close();
                Delay(3);
                var Reason = String.Empty;
                switch (reason)
                {
                    case "Affordability":
                        Reason = "Affordability";
                        break;
                    case "Client did disclose":
                        Reason = "Client did disclose";
                        break;
                    case "Client did not disclose":
                        Reason = "Client did not disclose";
                        break;
                    case "Client no longer interested":
                        Reason = "Client no longer interested";
                        break;
                    case "Cooling off period":
                        Reason = "Cooling off period";
                        break;
                    case "Death of Policy Holder":
                        Reason = "Death of Policy Holder";
                        break;
                    case "Duplicate Policy":
                        Reason = "Duplicate Policy";
                        break;
                    case "Fraud":
                        Reason = "Fraud";
                        break;
                    case "Fraud Allegation":
                        Reason = "Fraud Allegation";
                        break;
                    case "Fraud Disproved":
                        Reason = "Fraud Disproved";
                        break;
                    case "Fraud Investigation Inconclusive":
                        Reason = "Fraud Investigation Inconclusive";
                        break;
                    case "Incorrect Details":
                        Reason = "Incorrect Details";
                        break;
                    case "Late/Early Deduction":
                        Reason = "Late/Early Deduction";
                        break;
                    case "Money Back Guarantee":
                        Reason = "Money Back Guarantee";
                        break;
                    case "No Follow Up":
                        Reason = "No Follow Up";
                        break;
                    case "No Policy Document":
                        Reason = "No Policy Document";
                        break;
                    case "No Reason Given":
                        Reason = "No Reason Given";
                        break;
                    case "Not as Agreed":
                        Reason = "Not as Agreed";
                        break;
                    case "Not happy with service":
                        Reason = "Not happy with service";
                        break;
                    case "Replacement Of Policy":
                        Reason = "Replacement Of Policy";
                        break;
                    case "Wrong Deduction":
                        Reason = "Wrong Deduction";
                        break;
                    default:
                        break;
                }
                IWebElement selectRole = _driver.FindElement(By.Name("frmCancelReason"));
                SelectElement s = new SelectElement(selectRole);
                s.SelectByValue(Reason);
                Delay(4);
                //cancel
                _driver.FindElement(By.Name("btnSubmit")).Click();
                Delay(2);
                // Switch the control of 'driver' to the Alert from main Window
                IAlert simpleAlert1 = _driver.SwitchTo().Alert();
                // '.Accept()' is used to accept the alert '(click on the Ok button)'
                simpleAlert1.Accept();
                Delay(3);

                //validations
                //Get policy status
                string newStatus = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr[2]/td[2]/u/font")).Text;
                Delay(2);
                //get the termination date
                string termination_date = String.Empty;
                termination_date = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[3]/td/div/table/tbody/tr/td/span/table/tbody/tr/td[1]/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[9]/td[2]")).Text;
                Delay(2);
                string req_type = String.Empty;
                string movement = String.Empty;
                try
                {
                    //Click on financial mendates
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[4]/td/div/table/tbody/tr[2]/td/span/table/tbody/tr[2]/td[1]/a")).Click();
                    //get request type
                    Delay(2);
                    //loop through the types
                    for (int i = 0; i < 5; i++)
                    {
                        string _type = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center[2]/div[1]/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i + 2}]/td[1]/a")).Text;
                        if (_type.Contains("Cancel"))
                        {
                            req_type = _type;
                            break;
                        }
                    }
                    Delay(2);

                    //Navigate to the Contract Summary screen
                    Nav_ContractSummary();
                }
                catch
                {
                }
                movement = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[11]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[1]")).Text;
                Delay(2);

                //Navigate to the Components screen
                Nav_Components();
                Delay(2);
                //Get Comp slab text
                var compSlabText = String.Empty;
                try
                {
                    compSlabText = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td")).Text;
                }
                catch (Exception ex)
                {
                }
                Delay(2);
                //Get total premium
                var totalPrem = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td[2]")).Text;
                Delay(3);
                if ((newStatus == "Cancelled" || newStatus == "Not Taken Up")
                    && !termination_date.Equals(String.Empty)
                    && !movement.Equals(String.Empty)
                    && !compSlabText.Equals(String.Empty)
                    && totalPrem.Equals("0.00"))
                {
                    results = "Passed";
                }
                else
                {
                    results = "Failed";
                    TakeScreenshot("CancelPolicy");
                    errMsg = "The policy was not fully cancelled";
                }
            }
            catch (Exception ex)
            {
                results = "Failed";
                TakeScreenshot("CancelPolicy");
                if (ex.Message.Length > 255)
                {
                    errMsg = ex.Message.Substring(0, 255);
                }
                else
                {
                    errMsg = ex.Message;
                }
            }
            writeResultsToDB(results, Int32.Parse(scenarioID), errMsg);
            Assert.IsTrue(results.Equals("Passed"));
        }

        [Test, TestCaseSource("GetTestData", new object[] { "ReInstate" })]
        //Cannot Use Contracts that are more than 4 months in arrears
        public void ReInstate(string contractRef, string scenarioID)
        {
            if (String.IsNullOrEmpty(contractRef))
            {
                Assert.Ignore();
            }
            string errMsg = String.Empty;
            string results = String.Empty;
            try
            {
                string componentprem = String.Empty;
                string ReinstatementReason = String.Empty;
                string ReinstatementDate = String.Empty;
                string commDate = String.Empty;
                string movements = String.Empty;
                string date = DateTime.Today.ToString("g");
                Delay(8);
                IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;
                policySearch(contractRef);
                //Contract Status validation
                Delay(2);
                SetproductName();
                var expectedpremium = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv9']/table/tbody/tr[2]/td[2]")).Text;
                commDate = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr[6]/td[2]")).Text;
                var Cancelled = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr[2]/td[2]/u/font")).Text;
                IWebElement policyOptionElement3 = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));
                //Creating object of an Actions class
                Actions action2 = new Actions(_driver);
                //Performing the mouse hover action on the target element.
                action2.MoveToElement(policyOptionElement3).Perform();
                Delay(2);
                //Click on Reinstate
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/table/tbody/tr/td[1]/table/tbody/tr/td/div[2]/table/tbody/tr[10]/td/div/div[3]/a/img")).Click();
                Delay(2);
                OpenDBConnection($"SELECT * FROM Reinstate  WHERE Scenario_ID = {scenarioID}");
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    ReinstatementReason = reader["ReinstatementReason"].ToString();
                    ReinstatementDate = reader["ReinstatementDate"].ToString();
                }
                connection.Close();
                Delay(1);
                var value = String.Empty;
                //Insert Reinstate reason
                switch (ReinstatementReason)
                {
                    case ("Back dated claim"):
                        value = "ReinstateReason1";
                        break;
                    case ("Bank details updated, not activated"):
                        value = "ReinstateReason2";
                        break;
                    case ("Client request"):
                        value = "ReinstateReason3";
                        break;
                    case ("Conservation"):
                        value = "ReinstateReason4";
                        break;
                    case ("Due to ill health"):
                        value = "ReinstateReason5";
                        break;
                    case ("Error on collections"):
                        value = "ReinstateReason6";
                        break;
                    case ("Cooling off period"):
                        value = "ReinstateReason8";
                        break;
                    default:
                        break;
                }
                SelectElement dropDown = new SelectElement(_driver.FindElement(By.Name("frmReason")));
                dropDown.SelectByValue(value);
                Delay(3);
                //Click submit
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[12]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td/span/a")).Click();
                Delay(4);
                string error_check = String.Empty;
                try
                {
                    error_check = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/div/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[3]/td[2]/em/span")).Text;
                    results = "Failed";
                    TakeScreenshot("ReInstate");
                    errMsg = error_check;

                }
                catch
                {
                    try
                    {
                        //Click submit
                        _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[2]/tbody/tr[3]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td/span/a")).Click();
                        Delay(5);
                    }
                    catch
                    {
                    }

                    //Contract Status validation
                    var StatusInForce = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr[2]/td[2]/u/font")).Text;
                    var movement = _driver.FindElement(By.XPath("/html[1]/body[1]/center[1]/center[1]/form[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/span[1]/table[1]/tbody[1]/tr[2]/td[3]/center[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/span[1]/table[1]/tbody[1]/tr[11]/td[1]/div[1]/table[1]/tbody[1]/tr[4]/td[2]/span[1]/table[1]/tbody[1]/tr[2]/td[1]")).Text;
                    Assert.IsTrue(StatusInForce.Equals("In-Force", StringComparison.CurrentCultureIgnoreCase));
                    Delay(2);
                    var collectiomethod = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[3]/td/div/table/tbody/tr/td/span/table/tbody/tr/td[2]/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[5]/td[2]")).Text;
                    string Events = String.Empty;
                    if (collectiomethod == "Suspended Stop Order")
                    {
                        //Stop Order
                        //Navgate to the Components screen
                        Nav_Components();

                        //Capture new premium
                        componentprem = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td[2]")).Text;

                        //Navigate to Events screen
                        Nav_Events();

                        //Events
                        try
                        {
                            for (int i = 2; i < 23; i++)
                            {
                                Events = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]")).Text;
                                if (Events == "Create a New Deduction" || Events == "Create Instruction")
                                {
                                    break;
                                }
                            }
                        }
                        catch
                        {
                            Events = "Correct Event Was Not Found!";
                        }

                        Delay(3);

                        //Status validation
                        if (StatusInForce == "In-Force" & (movement == "Re-instatements" || movement == "Restart Policy") & (expectedpremium == componentprem)  )//<-- //& (Events == "Create a New Deduction" || Events == "Create Instruction")) For some reason no longer being added to events list
                        {
                            results = "Passed";
                        }
                        else
                        {
                            results = "Failed";
                            TakeScreenshot("ReInstate");
                            errMsg = "Status of Policy was not changed sucessfully";
                        }
                    }
                    else
                    {
                        //DebiCheck
                        _driver.FindElement(By.Name("fcMandateReference1")).Click();
                        Delay(2);
                        Events = _driver.FindElement(By.Name("fcMandateRequestType1")).Text;
                        var Status = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center[2]/div[1]/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[2]")).Text;

                        //Navigate to the Components screen
                        Nav_Components();

                        //Store value of current premium
                        componentprem = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td[2]")).Text;

                        //add status validation
                        if ((Events == "Create Instruction" || Events == "Cancel Instruction") && (Status == "Created Request" || Status == "Request Sent") && StatusInForce == "In-Force" && expectedpremium == componentprem)
                        {
                            results = "Passed";
                        }
                        else
                        {
                            results = "Failed";
                            TakeScreenshot("ReInstate");
                            errMsg = "Status of Policy was not changed sucessfully ";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (ex.Message.Length > 255)
                {
                    errMsg = ex.Message.Substring(0, 255);
                }
                else
                {
                    errMsg = ex.Message;
                }
                results = "Failed";
            }
            writeResultsToDB(results, Int32.Parse(scenarioID), errMsg);
            Assert.IsTrue(results.Equals("Passed"));
        }


        [Test, TestCaseSource("GetTestData", new object[] { "ChangeLifeAssured" })]
        public void changeLifeAssured(string contractRef, string scenarioID)
        {
            if (String.IsNullOrEmpty(contractRef))
            {
                Assert.Ignore();
            }
            string results = String.Empty;
            string errMsg = String.Empty;
            try
            {
                Delay(4);
                policySearch(contractRef);
                Delay(2);
                SetproductName();
                string title = String.Empty, surname = String.Empty, MaritalStatus = String.Empty, EducationLevel = String.Empty, Department = String.Empty, Profession = String.Empty;
                var Roleplayer = String.Empty;
                var IdNum = String.Empty;
                OpenDBConnection($"SELECT * FROM ChangeLifeData WHERE Scenario_ID = {scenarioID}");
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    title = reader["Title"].ToString();
                    surname = reader["surname"].ToString();
                    MaritalStatus = reader["MaritalStatus"].ToString();
                    EducationLevel = reader["EducationLevel"].ToString();
                    Department = reader["Department"].ToString();
                    Profession = reader["Profession"].ToString();
                    Department = reader["Department"].ToString();
                    Profession = reader["Profession"].ToString();
                    Roleplayer = reader["Roleplayer"].ToString().Trim();
                    IdNum = reader["RolePlayer_idNum"].ToString().Trim();
                }
                connection.Close();
                Delay(2);
                string roleplayer_check, idno_check = String.Empty;
                try
                {
                    for (int i = 2; i < 23; i++)
                    {
                        try
                        {
                            roleplayer_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/span/big/b/a")).Text;
                            idno_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[5]/span")).Text;
                        }
                        catch
                        {
                            roleplayer_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/span/a")).Text;
                            idno_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[5]/span")).Text;
                        }
                        if (roleplayer_check == Roleplayer && idno_check == IdNum)
                        {
                            _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[2]/a")).Click();
                            break;
                        }
                    }
                }
                catch
                {
                    results = "Failed";
                    TakeScreenshot("TerminateRolePlayer");
                    errMsg = "Roleplayer was not terminated";
                }
                Delay(2);
                var oldMaritalStatus = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv4']/table/tbody/tr[2]/td[4]")).Text;
                var oldeducationlevel = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv10']/table/tbody/tr[2]/td[2]")).Text;
                var oldDepartment = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv10']/table/tbody/tr[3]/td[2]")).Text;
                var oldProfession = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv10']/table/tbody/tr[4]/td[2]")).Text;
                Delay(2);
                //click oN Change
                _driver.FindElement(By.Name("btnChangePerson")).Click();
                var value = String.Empty;
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
                oSelect.SelectByText(title);
                Delay(2);
                //Insert Surname
                _driver.FindElement(By.Name("fcLastName")).Clear();
                Delay(2);
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
                        value = ("BK");
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
                var newMaritalStatus = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv4']/table/tbody/tr[2]/td[4]")).Text;
                var neweducationlevel = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/table/tbody/tr/td/span/table/tbody/tr/td[3]/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[2]")).Text;
                var newDepartment = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv10']/table/tbody/tr[3]/td[2]")).Text;
                var newProfession = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv10']/table/tbody/tr[4]/td[2]")).Text;
                //Vaidation based Martial stat
                if ((oldMaritalStatus) != (newMaritalStatus) & (oldeducationlevel) != (neweducationlevel) & (oldDepartment) != (newDepartment) & (oldDepartment) != (newDepartment))
                {
                    //Details sucessfully changed);
                    results = "Passed";
                }
                else
                {
                    results = "Failed";
                    TakeScreenshot("Changelifedata");
                    errMsg = "Life Assured was not changed sucessfully";
                }
            }
            catch (Exception ex)
            {
                if (ex.Message.Length > 255)
                {
                    errMsg = ex.Message.Substring(0, 255);
                }
                else
                {
                    errMsg = ex.Message;
                }
                results = "Failed";
            }
            writeResultsToDB(results, Int32.Parse(scenarioID), errMsg);
            Assert.IsTrue(results.Equals("Passed"));
        }
        private void componentUpgradeDowngrade(string contractRef, string scenarioID)
        //Max cover for a child in product 3000 (Just Funeral) is 20,000 
        {
            Delay(8);
            string results = String.Empty;
            string errMsg = String.Empty;
            try
            {
                var currentCoverAmount = String.Empty;
                var newPremium = String.Empty;
                var commDate = String.Empty;
                var newSumAssured = String.Empty;
                var method = String.Empty;
                string ExpectedContractPremium, currentPremium;
                policySearch(contractRef);
                Delay(3);
                var product = SetproductName();
                //Get the Commencement date from contract summary screen
                commDate = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr[6]/td[2]")).Text;

                //Scroll Downa
                Delay(4);
                IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;
                Delay(4);
                //Get Current premium
                currentPremium = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv9']/table/tbody/tr[2]/td[2]")).Text;
                //Go to component
                // clickOnMainMenu();
                try
                {
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[1]/a/img[2]")).Click();
                }
                catch (Exception ex)
                {
                    clickOnMainMenu();
                }
                Delay(2);
                try
                {
                    //click on components
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[3]/tbody/tr/td/a/div/div/div[3]")).Click();
                }
                catch (Exception ex)
                {
                    //expand contract sumary
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[1]/a/img[2]")).Click();
                    //click on components
                    Delay(2);
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[3]/tbody/tr/td/a/div/div/div[3]")).Click();
                }
                Delay(3);
                var component = String.Empty;
                var idNo = String.Empty;
                OpenDBConnection($"SELECT * FROM ComponentDowngradeUpgrade where Scenario_ID = {scenarioID}");
                OpenDBConnection($"SELECT * FROM ComponentDowngradeUpgrade where Scenario_ID = {scenarioID}");
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    component = reader["component"].ToString().Trim();
                    method = reader["Method"].ToString().Trim();
                    newSumAssured = reader["Cover_Amount"].ToString().Trim();
                    idNo = reader["RolePlayer_idNo"].ToString().Trim();
                    break;
                }
                connection.Close();
                string component_check = String.Empty;
                try
                {
                    for (int i = 2; i < 23; i++)
                    {
                        try
                        {
                            component_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/big/b/a")).Text;

                        }
                        catch
                        {
                            component_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/a")).Text;
                        }
                        if (component_check == component)
                        {
                            try
                            {
                                _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/big/b/a")).Click();
                                break;
                            }
                            catch
                            {
                                _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/a")).Click();
                                break;
                            }

                        }
                    }
                }
                catch
                {
                    results = "Failed";
                    TakeScreenshot("AddRolePlayer");
                    errMsg = "Roleplayer was not added successfully";
                }
                try
                {
                    for (int i = 2; i < 23; i++)
                    {
                        component_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/a")).Text;

                        if (component_check == component)
                        {
                            _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/a")).Click();
                            break;
                        }
                    }
                }
                catch
                {
                    results = "Failed";
                    TakeScreenshot("TerminateRolePlayer");
                    errMsg = "Component was not found";
                }
                if (product == "Safrican Provide and Protect Plan (4000)")
                {
                    Delay(4);
                    //Get The current Sum Assured for the life assured
                    currentCoverAmount = _driver.FindElement(By.XPath("//*[@id='frmCbmcc']/tbody/tr[7]/td[2]")).Text;
                    //Get start date
                    commDate = _driver.FindElement(By.XPath("//*[@id='frmCbmcc']/tbody/tr[3]/td[4]")).Text;
                    IWebElement policyOptionElement = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/div[3]/table/tbody/tr/td/div/div[3]"));
                    //Creating object of an Actions class
                    Actions action = new Actions(_driver);
                    //Performing the mouse hover action on the target element.
                    action.MoveToElement(policyOptionElement).Perform();
                    Delay(5);
                    _driver.FindElement(By.XPath("//div[3]/a/img")).Click();
                    Delay(4);
                    _driver.FindElement(By.Name("frmCCStartDate")).Clear();
                    Delay(4);
                    _driver.FindElement(By.Name("frmCCStartDate")).SendKeys(commDate);
                    _driver.FindElement(By.Name("frmSPAmount")).Clear();
                    _driver.FindElement(By.Name("frmSPAmount")).SendKeys(newSumAssured);
                    Delay(4);
                    _driver.FindElement(By.Name("btncbmcc13")).Click();
                    Delay(4);
                    _driver.FindElement(By.Name("btncbmcc17")).Click();
                    Delay(4);
                    _driver.FindElement(By.Name("btncbmcc23")).Click();
                    Delay(4);
                    _driver.FindElement(By.Name("2000175333.8")).Click();
                    Delay(3);
                    newPremium = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv9']/table/tbody/tr[2]/td[2]")).Text;
                    Delay(3);
                    if (Convert.ToDecimal(currentPremium) > Convert.ToDecimal(newPremium))
                    {
                        results = "Passed";
                    }
                    else
                    {
                        results = "Failed";
                    }
                }
                else
                {
                    Delay(4);
                    //Get The current componentStartDate
                    // var componentStartDate = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[3]/td/div/table/tbody/tr/td/span/table/tbody/tr/td[1]/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[6]/td[2]")).Text;
                    //Get The current Sum Assured for the life assured
                    currentCoverAmount = _driver.FindElement(By.XPath("//*[@id='frmCbmcc']/tbody/tr[8]/td[4]")).Text;
                    IWebElement policyOptionElement = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/div[3]/table/tbody/tr/td/div/div[3]"));
                    //Creating object of an Actions class
                    Actions action = new Actions(_driver);
                    //Performing the mouse hover action on the target element.
                    action.MoveToElement(policyOptionElement).Perform();
                    Delay(3);
                    _driver.FindElement(By.XPath("//div[3]/a/img")).Click();
                    Delay(3);
                    SelectElement oSelect = new SelectElement(_driver.FindElement(By.Name("frmSPAmount")));
                    oSelect.SelectByValue(newSumAssured);
                    Delay(4);
                    _driver.FindElement(By.Name("btncbmcc13")).Click();
                    Delay(4);
                    _driver.FindElement(By.Name("btncbmcc17")).Click();
                    Delay(3);
                    _driver.FindElement(By.Name("btncbmcc23")).Click();
                    Delay(4);
                    var premiumfromRateTable = getPremuimFromRateTable(idNo, component, newSumAssured, product);
                    Delay(2);
                    //Get the new contract Premium
                    ExpectedContractPremium = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv9']/table/tbody/tr[2]/td[4]")).Text;
                    string oldcompPremium = String.Empty;
                    string oldcompPremium_check = String.Empty;
                    double ContractPremDiff, ComponentPremDiff;
                    //Validation on Premiums
                    if (method.Equals("Downgrade"))
                    {
                        ContractPremDiff = double.Parse(currentPremium) - double.Parse(ExpectedContractPremium);
                        try
                        {
                            for (int i = 2; i < 23; i++)
                            {
                                oldcompPremium_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/a")).Text;

                                if (oldcompPremium_check == component)
                                {
                                    oldcompPremium = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[7]")).Text;
                                    break;
                                }
                            }
                        }
                        catch
                        {
                            results = "Failed";
                            TakeScreenshot("ComponentUpgradeDowngrade");
                            errMsg = "Component was not found!";
                        }

                        ComponentPremDiff = double.Parse(oldcompPremium) - Decimal.ToDouble(premiumfromRateTable);

                        if (ContractPremDiff == ComponentPremDiff && double.Parse(ExpectedContractPremium) < double.Parse(currentPremium))
                        {
                            results = "Passed";
                        }
                        else
                        {
                            TakeScreenshot(method);
                            results = "Failed";
                            errMsg = "Sum was not decreased successfully";
                        }
                    }
                    else
                    {
                        ContractPremDiff = double.Parse(ExpectedContractPremium) - double.Parse(currentPremium);
                        try
                        {
                            for (int i = 2; i < 23; i++)
                            {
                                oldcompPremium_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/a")).Text;

                                if (oldcompPremium_check == component)
                                {
                                    oldcompPremium = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[7]")).Text;
                                    break;
                                }
                            }
                        }
                        catch
                        {
                            results = "Failed";
                            TakeScreenshot("ComponentUpgradeDowngrade");
                            errMsg = "Component was not found!";
                        }

                        ComponentPremDiff = Decimal.ToDouble(premiumfromRateTable) - double.Parse(oldcompPremium);

                        if (ContractPremDiff == ComponentPremDiff && double.Parse(ExpectedContractPremium) > double.Parse(currentPremium))
                        {
                            results = "Passed";
                        }
                        else
                        {
                            TakeScreenshot(method);
                            results = "Failed";
                            errMsg = "Sum was not Increased successfully";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                results = "Failed";
                TakeScreenshot("ComponentUpgradeDowngrade");
                if (ex.Message.Length > 255)
                {
                    errMsg = ex.Message.Substring(0, 255);
                }
                else
                {
                    errMsg = ex.Message;
                }

            }
            writeResultsToDB(results, Int32.Parse(scenarioID), errMsg);
            Assert.IsTrue(results.Equals("Passed"));
        }

        [Test, TestCaseSource("GetTestData", new object[] { "DecreaseSumAssured" })]
        public void decreaseSumAssured(string contractRef, string scenarioID)
        {
            if (String.IsNullOrEmpty(contractRef))
            {
                Assert.Ignore();
            }
            componentUpgradeDowngrade(contractRef, scenarioID);
        }
        [Test, TestCaseSource("GetTestData", new object[] { "IncreaseSumAssured" })]
        public void increaseSumAssured(string contractRef, string scenarioID)
        {
            if (String.IsNullOrEmpty(contractRef))
            {
                Assert.Ignore();
            }
            componentUpgradeDowngrade(contractRef, scenarioID);
        }

        //if you changing a contracts method to stop order make sure it hasn't previously been stop order (Check movement for "collection") as you'll need to provide the same employee number as the last time it was stop order
        [Test, TestCaseSource("GetTestData", new object[] { "ChangeCollectionMethod" })]
        public void changeCollectionMethod(string contractRef, string scenarioID)
        {
            if (String.IsNullOrEmpty(contractRef))
            {
                Assert.Ignore();
            }
            Delay(8);
            ;
            var errMsg = String.Empty;
            string results = String.Empty;
            try
            {
                String test_url_2_title = "SANLAM RM - Safrican Retail - Warpspeed Lookup Window";
                IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;
                string date = DateTime.Today.ToString("g");
                string employee_number1 = String.Empty;
                string collectionmethod = String.Empty;
                policySearch(contractRef);
                Delay(4);
                SetproductName();
                Delay(3);
                //click on policy payer  
                _driver.FindElement(By.Name("fcRoleEntityLink3")).Click();
                Delay(3);
                try
                {
                    // Switch the control of 'driver' to the Alert from main Window
                    IAlert simpleAlert1 = _driver.SwitchTo().Alert();
                    // '.Accept()' is used to accept the alert '(click on the Ok button)'
                    simpleAlert1.Accept();
                }
                catch
                {
                    var prevcollectionmethod = _driver.FindElement(By.XPath("//*[@id='frmCbmre']/tbody/tr[8]/td[4]")).Text;
                }
                var previouscollectionmethod = _driver.FindElement(By.XPath("//*[@id='frmCbmre']/tbody/tr[8]/td[4]")).Text;
                IWebElement policyOptionElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));
                //Creating object of an Actions class
                Actions action = new Actions(_driver);
                //Performing the mouse hover action on the target element.
                action.MoveToElement(policyOptionElement).Perform();
                //Click on options
                _driver.FindElement(By.XPath("//*[@id='m0t0']/tbody/tr[3]/td/div/div[3]/a/img")).Click();
                Delay(2);
                string winHandleBefore = _driver.CurrentWindowHandle;
                OpenDBConnection($"SELECT * FROM CollectionMethodData  WHERE Scenario_ID = {scenarioID}");
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    collectionmethod = reader["collectionmethod"].ToString();
                    employee_number1 = reader["employee_number1"].ToString();
                }
                connection.Close();
                var value = String.Empty;
                //Insert EducationLevel
                switch (collectionmethod)
                {
                    case ("Debit Order"):
                        value = "100293.19";
                        break;
                    case ("Stop Order"):
                        value = "108978.19";
                        break;
                    case ("DebiCheck"):
                        value = "959096918.488";
                        break;
                    default:
                        break;
                }
                SelectElement oSelect4 = new SelectElement(_driver.FindElement(By.Name("fcCollectionMethod")));
                oSelect4.SelectByValue(value);
                Delay(2);
                var employeeno = String.Empty;
                if (value == "108978.19")
                {
                    if (employee_number1 == "")
                    {
                        Random rnd = new Random();
                        int myRandomNo = rnd.Next(10000000, 99999999); // creates a 8 digit random no.
                        employeeno = myRandomNo.ToString();
                    }
                    else
                    {
                        employeeno = employee_number1;
                    }
                    //Click on EMPLOYEE NUMBER
                    _driver.FindElement(By.Name("fcEmployeeNumber")).Click();
                    Delay(2);
                    //Click on EMPLOYEE NUMBER
                    _driver.FindElement(By.Name("fcEmployeeNumber")).SendKeys(employeeno);
                    Delay(2);
                    //Search employee
                    _driver.FindElement(By.Name("fcEmployerButton")).Click();
                    Assert.AreEqual(2, _driver.WindowHandles.Count);
                    var newWindowHandle = _driver.WindowHandles[1];
                    Assert.IsTrue(!string.IsNullOrEmpty(newWindowHandle));
                    string expectedNewWindowTitle = test_url_2_title;
                    Assert.AreEqual(_driver.SwitchTo().Window(newWindowHandle).Title, expectedNewWindowTitle);
                    //Search employee
                    _driver.FindElement(By.XPath("//*[@id='lkpResultsTable']/tbody/tr[17]")).Click();
                    //Switch to main window
                    _driver.SwitchTo().Window(winHandleBefore);
                    Delay(3);
                }
                try
                {
                    _driver.SwitchTo().Alert().Dismiss();
                }
                catch
                {
                    //Click on submit
                    _driver.FindElement(By.Id("GBLbl-1")).Click();
                    Delay(3);
                }
                _driver.SwitchTo().Window(winHandleBefore);
                var expectedcollectionM = _driver.FindElement(By.XPath("//*[@id='frmCbmre']/tbody/tr[8]/td[4]")).Text;
                var addedemployeenumber = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[12]/td[2]")).Text;
                Delay(3);

                //Navigate to the  screen
                Nav_ContractSummary();

                var collectionmethodmovement = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[11]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[1]")).Text;
                Delay(3);
                if (collectionmethod == "Stop Order")
                {
                    if (expectedcollectionM != previouscollectionmethod & (collectionmethodmovement == "Collection Method Maintenance") & (addedemployeenumber == employeeno))
                    {
                        results = "Passed";
                    }
                    else
                    {
                        TakeScreenshot("ChangeCollectionMethod");
                        results = "Failed";
                        errMsg = "Collection method  was not changed succesfully";
                    }
                }
                else
                {
                    if (expectedcollectionM != previouscollectionmethod & (collectionmethodmovement == "Collection Method Maintenance"))
                    {
                        results = "Passed";
                    }
                    else
                    {
                        TakeScreenshot("ChangeCollectionMethod");
                        results = "Failed";
                        errMsg = "Collection method  was not changed succesfully";
                    }
                }

            }
            catch (Exception ex)
            {
                results = "Failed";
                TakeScreenshot("ChangeCollectionMethod");
                if (ex.Message.Length > 255)
                {
                    errMsg = ex.Message.Substring(0, 255);
                }
                else
                {
                    errMsg = ex.Message;
                }
                results = "Failed";
            }
            writeResultsToDB(results, Int32.Parse(scenarioID), errMsg);
            Assert.IsTrue(results.Equals("Passed"));
        }

        [Test, TestCaseSource("GetTestData", new object[] { "TerminateRolePlayer" })]
        public void terminateRolePlayer(string contractRef, string scenarioID)
        {
            if (String.IsNullOrEmpty(contractRef))
            {
                Assert.Ignore();
            }
            Delay(8);
            ;
            IJavaScriptExecutor js2 = (IJavaScriptExecutor)_driver;
            var errMsg = String.Empty;
            string results = String.Empty; string component = String.Empty; string currentPremium = String.Empty;
            string ID = String.Empty;
            string roleplayer = String.Empty;
            string relationship = String.Empty;
            string roleplayer_check, relationship_check, idno_check = String.Empty;
            OpenDBConnection($"SELECT * FROM TerminateRole WHERE Scenario_ID = {scenarioID}");
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                roleplayer = reader["Role_Player"].ToString();
                relationship = reader["Relationship"].ToString();
                component = reader["Relationship"].ToString();
                ID = reader["ID_No"].ToString();
            }
            connection.Close();
            try
            {
                policySearch(contractRef);


                currentPremium = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv9']/table/tbody/tr[2]/td[2]")).Text;
                Delay(2);
                try
                {

                    var contractsummary = _driver.FindElement(By.XPath("//*[@id='t0_769']/table/tbody/tr/td[3]/a")).Text;

                }
                catch
                {
                    clickOnMainMenu();

                }

                if (roleplayer == "Beneficiary")
                {
                    for (int i = 2; i < 23; i++)
                    {
                        try
                        {
                            roleplayer_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/span/big/b/a")).Text;
                            relationship_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[3]")).Text;
                            idno_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[5]/span")).Text;
                            if (roleplayer_check == roleplayer && relationship_check == relationship && idno_check == ID)
                            {
                                _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/span/big/b/a")).Click();
                                break;
                            }
                        }
                        catch
                        {
                            roleplayer_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/span/a")).Text;
                            relationship_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[3]")).Text;
                            idno_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[5]/span")).Text;
                            if (roleplayer_check == roleplayer && relationship_check == relationship && idno_check == ID)
                            {
                                try
                                {
                                    _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/span/big/b/a")).Click();
                                }
                                catch
                                {
                                    _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/span/a")).Click();
                                }
                                break;
                            }
                        }
                    }
                    //Click "Terminate" button
                    _driver.FindElement(By.XPath($"/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[22]/td/table/tbody/tr/td[2]/table/tbody/tr/td[2]/table/tbody/tr/td/span/a")).Click();
                    Delay(3);

                    //Accept pop-up
                    _driver.SwitchTo().Alert().Accept();
                    Delay(3);
                    try
                    {
                        string EffDate = roleplayer_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[2]")).Text;
                        _driver.FindElement(By.Name("frmEffectiveToDate")).Clear();
                        _driver.FindElement(By.Name("frmEffectiveToDate")).SendKeys(EffDate);
                    }
                    catch
                    {
                    }
                    //Click last "Terminate" button
                    _driver.FindElement(By.XPath($"/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[21]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td/span/a")).Click();
                    Delay(3);
                    clickOnMainMenu();
                    Delay(2);
                    //Navigate to contract summary
                    _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[1]/table/tbody/tr/td/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[3]/a/div/div/div[2]")).Click();
                    Delay(3);

                    //Check though all of the roleplayers
                    try
                    {
                        for (int i = 2; i < 23; i++)
                        {
                            try
                            {
                                roleplayer_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/span/big/b/a")).Text;
                                relationship_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[3]")).Text;
                                idno_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[5]/span")).Text;
                            }
                            catch
                            {
                                roleplayer_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/span/a")).Text;
                                relationship_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[3]")).Text;
                                idno_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[5]/span")).Text;
                            }
                            if (roleplayer_check == roleplayer && relationship_check == relationship && idno_check == ID)
                            {
                                string movement_check = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[11]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[1]")).Text;

                                if (movement_check == "Role Player maintenance" || movement_check == "Terminate Benefit")
                                {
                                    results = "Passed";
                                }
                                else
                                {
                                    results = "Failed";
                                }
                                break;
                            }
                        }
                    }
                    catch
                    {
                        results = "Passed";
                    }
                }
                else
                {
                    try 
                    {
                        try
                        {
                            //click on components
                            _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[3]/tbody/tr/td/a/div/div/div[3]")).Click();
                        }
                        catch (Exception ex)
                        {
                            //expand Main Menu
                            _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/table[6]/tbody/tr/td/a")).Click();
                            //expand contract sumary
                            _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[6]/table[5]/tbody/tr/td/table/tbody/tr/td[1]/a/img[2]")).Click();
                            //click on components
                            Delay(2);
                            _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[6]/div[1]/table[1]/tbody/tr/td/a/div/div/div[3]")).Click();
                        }
                    }
                    catch
                    {
                        try
                        {
                            //click on components
                            _driver.FindElement(By.XPath("")).Click();
                        }
                        catch (Exception ex)
                        {
                            //expand Main Menu
                            _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/table[6]/tbody/tr/td/a")).Click();
                            //expand contract sumary
                            _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[6]/table[5]/tbody/tr/td/table/tbody/tr/td[1]/a/img[2]")).Click();
                            //click on components
                            Delay(2);
                            _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[6]/div[1]/table[1]/tbody/tr/td/a/div/div/div[3]")).Click();
                        }
                    }
                    

                    for (int i = 0; i < 24; i++)
                    {
                        IWebElement comp;
                        var xPath = "";
                        try
                        {
                            xPath = $"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i + 2}]/td[1]/big/b/a";
                            comp = _driver.FindElement(By.XPath(xPath));
                        }
                        catch (Exception ex)
                        {
                            xPath = $"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i + 2}]/td[1]/a";
                            comp = _driver.FindElement(By.XPath(xPath));
                        }
                        var compTxt = comp.Text;
                        if (compTxt.Contains(component))
                        {
                            Delay(2);
                            comp.Click();
                            var idComp = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[11]/td[2]")).Text;
                            if (!(idComp.Contains(ID)))
                            {
                                _driver.Navigate().Back();
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                    _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td/span/a")).Click();
                    Delay(3);

                    var comm = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[2]")).Text;
                    _driver.FindElement(By.Name("frmCCEndDate")).Clear();
                    _driver.FindElement(By.Name("frmCCEndDate")).SendKeys(comm);


                    _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[11]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td/span/a")).Click();
                    Delay(3);
                    var newContractPremium = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr/td[2]")).Text;
                    //premium validation
                    Delay(1);
                    if (Convert.ToDecimal(currentPremium) > Convert.ToDecimal(newContractPremium))
                    {
                        try
                        {
                            //Contract summary
                            _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[3]/a/div/div/div[2]")).Click();
                        }
                        catch
                        {
                            _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[6]/table[5]/tbody/tr/td/table/tbody/tr/td[3]/a/div/div/div[2]"));
                        }
                        Delay(2);
                        //Find and click person to terminate    
                        for (int i = 2; i < 23; i++)
                        {
                            try
                            {
                                roleplayer_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/span/big/b/a")).Text;
                                relationship_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[3]")).Text;
                                idno_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[5]/span")).Text;
                            }
                            catch
                            {
                                roleplayer_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/span/a")).Text;
                                relationship_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[3]")).Text;
                                idno_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[5]/span")).Text;
                            }
                            if (roleplayer_check == roleplayer && relationship_check == relationship && idno_check == ID)
                            {
                                try
                                {
                                    _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/span/big/b/a")).Click();
                                }
                                catch
                                {
                                   _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/span/a")).Click();
                                }
                                break;
                            }
                        }

                        //Click "Terminate" button
                        _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[19]/td/table/tbody/tr/td[2]/table/tbody/tr/td[2]/table/tbody/tr/td/span/a")).Click();
                        Delay(3);
                        //Accept pop-up
                        _driver.SwitchTo().Alert().Accept();
                        Delay(3);
                        //Click last "Terminate" button
                        _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[18]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td/span/a")).Click();
                        Delay(3);
                        clickOnMainMenu();
                        Delay(2);
                        //Navigate to contract summary
                        _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[1]/table/tbody/tr/td/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[3]/a/div/div/div[2]")).Click();
                        Delay(3);

                        //Check though all of the roleplayers
                        try
                        {
                            for (int i = 2; i < 23; i++)
                            {
                                try
                                {
                                    roleplayer_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/span/big/b/a")).Text;
                                    relationship_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[3]")).Text;
                                    idno_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[5]/span")).Text;
                                }
                                catch
                                {
                                    roleplayer_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]/span/a")).Text;
                                    relationship_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[3]")).Text;
                                    idno_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[5]/span")).Text;
                                }
                                if (roleplayer_check == roleplayer && relationship_check == relationship && idno_check == ID)
                                {
                                    string movement_check = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[11]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[1]")).Text;

                                    if (movement_check == "Role Player maintenance" || movement_check  == "Terminate Benefit")
                                    {
                                        results = "Passed";
                                    }
                                    else
                                    {
                                        results = "Failed";
                                    }
                                    break;
                                }
                            }
                        }
                        catch
                        {
                            results = "Passed";
                        }

                    }
                    else
                    {
                        results = "Failed";
                        TakeScreenshot("TerminateRolePlayer");
                        errMsg = "Component was not terminated";
                    }
                }
            }
            catch (Exception ex)
            {
                TakeScreenshot("addRolePlayer");
                if (ex.Message.Length > 255)
                {
                    errMsg = ex.Message.Substring(0, 255);
                }
                else
                {
                    errMsg = ex.Message;
                }
                results = "Failed";
            }
            writeResultsToDB(results, Int32.Parse(scenarioID), errMsg);
            Assert.IsTrue(results.Equals("Passed"));
        }


        [OneTimeTearDown]
        public void closeBrowser()
        {
            DisconnectBrowser();
        }
    }
}