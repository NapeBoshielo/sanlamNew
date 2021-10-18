using NUnit.Framework;

using static ILR_TestSuite.TestBase;

using OpenQA.Selenium;

using OpenQA.Selenium.Interactions;

using System;

using System.Collections.Generic;
using ILR_TestSuite;
using System.Linq;

namespace PolicyServicing

{

    [TestFixture]

    public class PolicyServicing : TestBase
    {



        private IWebDriver _driver;

        [SetUp]

        public void startBrowser()

        {



            _driver = base.SiteConnection();

        }



        [Test, Order(1)]

        public void PolicyServicingTestSuite()

        {

            Delay(2);


            //LogInValidation();
            Add();
            Terminate();
            Add_PostDated();
            TerminatePostDated();
            Upgrades();
            Downgrades();
            UpgradesPostDated();
            DowngradesPostDated();
            MultiplePostDatedAlterations();
            CancelPolicy();
            PinChanges();
            AddBeneficiary();
            ChangeBeneficiary();
            AddPersonDetails();
            ChangePersonDetails();
            BankingDetailsPremiumPayer();
            CollectionMethod();
            ChangeEscalation();
            CancelEscalation();
            Reinstatement();



            Delay(20);

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

        [Category("Search")]


        public void SearchByProductName(string product)
        {

           


        //Click on product
        _driver.FindElement(By.CssSelector("")).Click();
            Delay(2);

            //Click on contract search button
            _driver.FindElement(By.LinkText("Search")).Click();
            Delay(2);

            //Click on contract number
            _driver.FindElement(By.LinkText("6000789")).Click();
            Delay(2);


            //Click on Add Roleplayer
            _driver.FindElement(By.Name("btnAddRolePlayer")).Click();
            Delay(2);


            //Click on Role
            _driver.FindElement(By.LinkText("frmRoleObj")).Click();
            Delay(2);

            _driver.FindElement(By.LinkText("frmRoleObj")).Click();
            Delay(2);

        }


        [Category("Add")]
        public void Add()
        {

            try

            {
                Delay(2);
                //Click on the Insurance Tab
                _driver.FindElement(By.Name("PF_User_Menu")).Click();
                Delay(2);


                //Click on the Contract Search 
                _driver.FindElement(By.Name("cb_Broker_cbcts")).Click();
                Delay(2);


                //Click on Search tab
                _driver.FindElement(By.Name("frmProductObjLkpImg")).Click();
                Delay(2);

                SearchByProductName("Legacy Saver");


               





                ////Click on contract number
                //_driver.FindElement(By.LinkText("6000789")).Click();
                //Delay(2);









                //Select  New Business 


                //Select the Cotract type


                //Search the contract you want to Downgrade and Select it 




            }

            catch (Exception ex)

            {

                DisconnectBrowser();

                throw ex;

            }

        }

        [Category("Terminate")]

        public void Terminate()

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





        [Category("Add_PostDated")]

        public void Add_PostDated()

        {

            try

            {

            
            }

            catch (Exception ex)

            {
                base.TakeScreenshot(_driver,"","FailNewBusinessTest");

                DisconnectBrowser();

                throw ex;

            }

        }





  


        [Category("TerminatePostDated")]

        public void TerminatePostDated()

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

        [Category("Upgrades")]

        public void Upgrades()

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

        [Category("Downgrades")]

        public void Downgrades()

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

        [Category("UpgradesPostDated")]

        public void UpgradesPostDated()

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

        [Category("DowngradesPostDated")]

        public void DowngradesPostDated()

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

        [Category("MultiplePostDatedAlterations")]

        public void MultiplePostDatedAlterations()

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

        [Category("CancelPolicy")]

        public void CancelPolicy()

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

        
        [Category("PinChanges")]

        public void PinChanges()

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

        [Category("AddBeneficiary")]

        public void AddBeneficiary()

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
        [Category("ChangeBeneficiary")]

        public void ChangeBeneficiary()

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
        [Category("AddPersonDetails")]

        public void AddPersonDetails()

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

        [Category("ChangePersonDetails")]

        public void ChangePersonDetails()

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


        [Category("BankingDetailsPremiumPayer")]

        public void BankingDetailsPremiumPayer()

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

        [Category("CollectionMethod")]

        public void CollectionMethod()

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
        [Category("ChangeEscalation")]

        public void ChangeEscalation()

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
        [Category("CancelEscalation")]

        public void CancelEscalation()

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

        [Category("Reinstatement")]

        public void Reinstatement()

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

        [TearDown]

        public void closeBrowser()

        {

            base.DisconnectBrowser();

        }



    }

}
