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

    public class NewBusiness : TestBase
    {



        private IWebDriver _driver;

        [SetUp]

        public void startBrowser()

        {



            _driver = base.SiteConnection();

        }



        [Test, Order(1)]

        public void NewBusinessTestSuite()

        {

            Delay(2);


            LogInValidation();

            CreateQuote();

            QuoteAcceptance();

            CancelQuote();

            




            Delay(20);

        }


        [Category("LogInValidation")]

        public void LogInValidation()

        {

            try

            {



                TakeScreenshot(_driver, $@"{_screenShotFolder}\Login\", "LoginValidation");



            }

            catch (Exception ex)

            {

                DisconnectBrowser();

                throw ex;

            }

        }

        [Category("CreateQuote")]
        public void CreateQuote()

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

        [Category("QuoteAcceptance")]

        public void QuoteAcceptance()

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





        [Category("CancelQuote")]

        public void CancelQuote()

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


        [TearDown]

        public void closeBrowser()

        {

            base.DisconnectBrowser();

        }



    }

}
