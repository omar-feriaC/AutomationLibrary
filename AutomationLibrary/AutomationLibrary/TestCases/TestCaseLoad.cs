using NUnit.Framework;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationLibrary.TestCases
{
    [TestFixture]
    class TestCaseLoad
    {
        public bool blStop;

        [OneTimeSetUp]
        public void BeforeClass() 
        {
            blStop = ClsReportResult.fnExtentSetup();
            if (!blStop)
                AfterClass();
        }

        
        public void SetUp(string pstrTestCase) 
        {
            ClsReportResult.objTest = ClsReportResult.objExtent.CreateTest(pstrTestCase);
            ClsWebBrowser.fnInitBrowser("Chrome");
        }

        [Test]
        public void Test()
        {
            SetUp("Test");
            ClsWebBrowser.fnNavigateToUrl("https://intake-uat.sedgwick.com/");
            if (ClsWebElements.fnElementExist("Cookies button", By.Id("cookie-accept"), false)) 
            {
                ClsWebElements.fnClick(ClsWebElements.fnGetWebElement(By.Id("cookie-accept")), "cookie", false);
            }
            ClsWebElements.fnSendKeys(ClsWebElements.fnGetWebElement(By.Id("orangeForm-name")), "dsfsdf", "text", false);
            ClsWebElements.fnSendKeys(ClsWebElements.fnGetWebElement(By.Id("orangeForm-pass")), "dsfsdfsfd", "tex13123t", false);
            ClsWebElements.fnSendKeys(ClsWebElements.fnGetWebElement(By.Id("orangeForm-pass222")), "dfsdfdsfsfd", "tex13123t", false);
            TearDown();
        }


        public void TearDown() 
        {
            ClsWebBrowser.fnCloseBrowser();
        }

        [OneTimeTearDown]
        public void AfterClass() 
        {
            try
            {
                ClsReportResult.objExtent.Flush();
            }
            catch (Exception objException)
            {
                Console.WriteLine(objException.Message);
            }
        }

    }
}
