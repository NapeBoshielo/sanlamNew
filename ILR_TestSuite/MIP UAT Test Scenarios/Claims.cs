using NUnit.Framework;

using static ILR_TestSuite.TestBase;

using OpenQA.Selenium;

using OpenQA.Selenium.Interactions;

using System;

using System.Collections.Generic;
using ILR_TestSuite;

namespace ILR_TestSuite.MIP_UAT_Test_Scenarios

{

    [TestFixture]

    public class Claims : TestBase
    {



        private IWebDriver _driver;

        [SetUp]

        public void startBrowser()

        {



            _driver = base.SiteConnection();

        }



        [Test, Order(1)]

        public void ClaimsTestSuite()

        {

            Delay(2);


            LogInValidation();

            DownGrade();

            Upgrade();

            NewBusiness();

            Termination();




            Delay(20);

        }


        [Category("LogInValidation")]

        public void LogInValidation()

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

        [Category("DownGrade")]
        public void DownGrade()

        {

            try

            {
                //Click on the Insurance Tab



                //Click on the Contract Search 




                //Select  New Business 


                //Select the Cotract type


                //Search the contract you want to Downgrade and Select it 

                //






            }

            catch (Exception ex)

            {

                DisconnectBrowser();

                throw ex;

            }

        }

        [Category("Upgrade")]

        public void Upgrade()

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





        [Category("NewBusiness")]

        public void NewBusiness()

        {

            try

            {


            }

            catch (Exception ex)

            {
                base.TakeScreenshot(_driver, "", "FailNewBusinessTest");

                DisconnectBrowser();

                throw ex;

            }

        }








        [Category("Termination")]

        public void Termination()

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

