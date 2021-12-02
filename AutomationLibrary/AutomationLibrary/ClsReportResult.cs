using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationLibrary
{
    public static class ClsReportResult
    {
        
        public static ExtentReports objExtent;
        public static ExtentTest objTest;
        public static ExtentV3HtmlReporter objHtmlReporter;
        public static bool TC_Status;

        /// <summary>
        /// Setup tjhe intance of extent reports object
        /// </summary>
        /// <returns></returns>
        public static bool fnExtentSetup()
        {
            bool blSuccess;
            try
            {
                ClsDataDriven clsDD = new ClsDataDriven();
                blSuccess = clsDD.fnAutomationSettings();

                if (blSuccess)
                {
                    //To create report directory and add HTML report into it
                    objHtmlReporter = new ExtentV3HtmlReporter(ClsDataDriven.strReportLocation + ClsDataDriven.strReportName + @"\" + ClsDataDriven.strReportName + ".html");

                    objHtmlReporter.Config.ReportName = ClsDataDriven.strReportName;
                    objHtmlReporter.Config.DocumentTitle = ClsDataDriven.strProjectName + " - " + ClsDataDriven.strReportName;
                    objHtmlReporter.Config.Theme = AventStack.ExtentReports.Reporter.Configuration.Theme.Dark;
                    objHtmlReporter.Config.Encoding = "utf-8";

                    objExtent = new ExtentReports();
                    objExtent.AddSystemInfo("Project", ClsDataDriven.strProjectName);
                    objExtent.AddSystemInfo("Browser", ClsDataDriven.strBrowser);
                    objExtent.AddSystemInfo("Env", ClsDataDriven.strReportEnv);
                    objExtent.AttachReporter(objHtmlReporter);
                }
            }
            catch (Exception pobjException)
            {
                throw (pobjException);
            }
            return blSuccess;
        }

        /// <summary>
        /// Close and generates the HTML Extent Report
        /// </summary>
        public static void fnExtentClose()
        {
            try
            {
                var objStatus = TestContext.CurrentContext.Result.Outcome.Status;
                var objStacktrace = "" + TestContext.CurrentContext.Result.StackTrace + "";
                var objErrorMessage = TestContext.CurrentContext.Result.Message;
                Status objLogstatus;
                switch (objStatus)
                {
                    case TestStatus.Failed:
                        objLogstatus = Status.Fail;
                        TestContext.Progress.WriteLine($"Test ended with {objLogstatus} – {objErrorMessage}");
                        objTest.Log(objLogstatus, "Test ended with " + objLogstatus + " – " + objErrorMessage);
                        break;
                    case TestStatus.Passed:
                        objLogstatus = Status.Pass;
                        break;
                    case TestStatus.Inconclusive:
                        objLogstatus = Status.Pass;
                        break;
                    default:
                        objLogstatus = Status.Warning;
                        TestContext.Progress.WriteLine($"The status: {objLogstatus} is not supported.");
                        Console.WriteLine("The status: " + objLogstatus + " is not supported.");
                        break;
                }
            }
            catch (Exception pobjException)
            {
                throw (pobjException);
            }

        }

        /// <summary>
        /// Create a log step and optional takes the screenshot
        /// </summary>
        /// <param name="pstrStepName"></param>
        /// <param name="pstrDescription"></param>
        /// <param name="pstrStatus"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        [Obsolete("New function replace pstrStatus parameter to receive Status.[your status]")]
        public static void fnLog(string pstrStepName, string pstrDescription, string pstrStatus, bool pblScreenShot, bool pblHardStop = false, string pstrHardStopMsg = "")
        {
            if (pblScreenShot)
            {
                string strSCLocation = fnGetScreenshot();
                switch (pstrStatus.ToUpper())
                {
                    case "PASS":
                        TestContext.Progress.WriteLine($"{pstrDescription} - Pass");
                        objTest.Log(Status.Pass, pstrDescription, MediaEntityBuilder.CreateScreenCaptureFromPath(strSCLocation).Build());
                        break;
                    case "FAIL":
                        TC_Status = false;
                        TestContext.Progress.WriteLine($"{pstrDescription} - Fail");
                        objTest.Log(Status.Fail, pstrDescription, MediaEntityBuilder.CreateScreenCaptureFromPath(strSCLocation).Build());
                        if (pblHardStop)
                            Assert.Fail(pstrHardStopMsg);
                        break;
                    case "INFO":
                        TestContext.Progress.WriteLine($"{pstrDescription} - Info");
                        objTest.Log(Status.Info, pstrDescription, MediaEntityBuilder.CreateScreenCaptureFromPath(strSCLocation).Build());
                        break;
                    case "WARNING":
                        TestContext.Progress.WriteLine($"{pstrDescription} - Warning");
                        objTest.Log(Status.Warning, pstrDescription, MediaEntityBuilder.CreateScreenCaptureFromPath(strSCLocation).Build());
                        break;
                }
            }
            else
            {
                switch (pstrStatus.ToUpper())
                {
                    case "PASS":
                        TestContext.Progress.WriteLine($"{pstrDescription} - Pass");
                        objTest.Log(Status.Pass, pstrDescription);
                        break;
                    case "FAIL":
                        TC_Status = false;
                        TestContext.Progress.WriteLine($"{pstrDescription} - Fail");
                        objTest.Log(Status.Fail, pstrDescription);
                        if (pblHardStop) { Assert.Fail(pstrHardStopMsg); }
                        break;
                    case "INFO":
                        TestContext.Progress.WriteLine($"{pstrDescription} - Info");
                        objTest.Log(Status.Info, pstrDescription);
                        break;
                    case "WARNING":
                        TestContext.Progress.WriteLine($"{pstrDescription} - Warning");
                        objTest.Log(Status.Info, pstrDescription);
                        break;
                }
            }
        }

        /// <summary>
        /// Create a log step and optional takes the screenshot
        /// </summary>
        /// <param name="pstrStepName"></param>
        /// <param name="pstrDescription"></param>
        /// <param name="pstrStatus"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        [Obsolete("This method will be deleted soon, use fnLog(string pstrStepName, string pstrDescription, Status pstrStatus, bool pblScreenShot) instead")]
        public static void fnLog(string pstrStepName, string pstrDescription, Status pstrStatus, bool pblScreenShot, bool pblHardStop = false, string pstrHardStopMsg = "")
        {
            fnLog(pstrStepName, pstrDescription, pstrStatus, pblScreenShot);
        }

        /// <summary>
        /// Create a log step and optional takes the screenshot
        /// </summary>
        /// <param name="pstrStepName"></param>
        /// <param name="pstrDescription"></param>
        /// <param name="pstrStatus"></param>
        /// <param name="pblScreenShot"></param>
        public static void fnLog(string pstrStepName, string pstrDescription, Status pstrStatus, bool pblScreenShot)
        {
            MediaEntityModelProvider ss = null;
            if (pblScreenShot)
            {
                string strSCLocation = fnGetScreenshot();
                ss = MediaEntityBuilder.CreateScreenCaptureFromPath(strSCLocation).Build();
            }

            TestContext.Progress.WriteLine($"{pstrStepName}: {pstrDescription}");
            objTest.Log(pstrStatus, pstrDescription, ss);
        }

        /// <summary>
        /// Fails the current test scenario by closing the browser and throwing an exception
        /// </summary>
        /// <param name="pstrMessage">Fail message</param>
        public static void fnAssertFail(string pstrMessage)
        {
            TC_Status = false;
            ClsWebBrowser.fnCloseBrowser();
            throw new Exception($"Assertion Failed: {pstrMessage}");
        }

        /// <summary>
        /// Takes a screenshot during  execution time
        /// </summary>
        /// <returns></returns>
        public static string fnGetScreenshot()
        {
            string strSCName = "SC_" + ClsDataDriven.strProjectName + "_" + DateTime.Now.ToString("MMddyyyy_hhmmss");

            //To take screenshot
            Screenshot objFile = ((ITakesScreenshot)ClsWebBrowser.objDriver).GetScreenshot();

            string strFileLocation = ClsDataDriven.strReportLocation + @"\Screenshots\" + strSCName + ".jpg";
            //To save screenshot
            objFile.SaveAsFile(strFileLocation, ScreenshotImageFormat.Jpeg);

            return strFileLocation;
        }





    }
}
