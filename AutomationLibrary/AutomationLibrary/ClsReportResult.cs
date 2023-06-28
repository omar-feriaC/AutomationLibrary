using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
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
        public static bool DevOpsResult;
        public static bool isWarning;
        public static string BaseReportFolder;
        public static string FullReportFolder;
        public static string TestPlanSuite = "";
        
        
        /// <summary>
        /// Setup tjhe intance of extent reports object
        /// </summary>
        /// <returns></returns>
        public static bool fnExtentSetup()
        {
            bool blSuccess;
            try
            {
                //ClsDataDriven clsDD = new ClsDataDriven();
                ///blSuccess = clsDD.fnAutomationSettings();
                blSuccess = true;

                if (blSuccess)
                {
                    //To create report directory and add HTML report into it
                    //objHtmlReporter = new ExtentV3HtmlReporter(ClsDataDriven.strReportLocation + ClsDataDriven.strReportName + @"\" + ClsDataDriven.strReportName + ".html");
                    string folderID = $"_{Environment.UserName.ToString()}";
                    BaseReportFolder = TestContext.Parameters["GI_ReportLocation"] + DateTime.Now.ToString("MMddyyyy_hhmmss") + folderID + @"\";
                    FullReportFolder = BaseReportFolder + TestContext.Parameters["GI_ProjectName"] + @"\" + TestContext.Parameters["GI_ReportName"] + ".html";
                    fnFolderSetup();

                    objHtmlReporter = new ExtentV3HtmlReporter(FullReportFolder);


                    /*
                    objHtmlReporter.Config.ReportName = ClsDataDriven.strReportName;
                    objHtmlReporter.Config.DocumentTitle = ClsDataDriven.strProjectName + " - " + ClsDataDriven.strReportName;
                    objHtmlReporter.Config.Theme = AventStack.ExtentReports.Reporter.Configuration.Theme.Dark;
                    objHtmlReporter.Config.Encoding = "utf-8";
                    */

                    objHtmlReporter.Config.ReportName = TestContext.Parameters["GI_ReportName"];
                    objHtmlReporter.Config.DocumentTitle = TestContext.Parameters["GI_ProjectName"] + " - " + TestContext.Parameters["GI_ReportName"];
                    objHtmlReporter.Config.Theme = AventStack.ExtentReports.Reporter.Configuration.Theme.Dark;
                    objHtmlReporter.Config.Encoding = "utf-8";


                    /*
                    objExtent = new ExtentReports();
                    objExtent.AddSystemInfo("Project", ClsDataDriven.strProjectName);
                    objExtent.AddSystemInfo("Browser", ClsDataDriven.strBrowser);
                    objExtent.AddSystemInfo("Env", ClsDataDriven.strReportEnv);
                    objExtent.AttachReporter(objHtmlReporter);
                    */

                    objExtent = new ExtentReports();
                    objExtent.AddSystemInfo("Project", TestContext.Parameters["GI_ProjectName"]);
                    objExtent.AddSystemInfo("Browser", TestContext.Parameters["GI_BrowserName"]);
                    objExtent.AddSystemInfo("Env", TestContext.Parameters["GI_TestEnvironment"]);
                    objExtent.AddSystemInfo("Executed By", Environment.UserName.ToString());
                    objExtent.AddSystemInfo("Executed Machine", Environment.MachineName.ToString());
                    objExtent.AddSystemInfo("Execution Time", DateTime.Now.ToString("MM/ddy/yyy hh:mm:ss"));
                    objExtent.AttachReporter(objHtmlReporter);
                }
            }
            catch (Exception pobjException)
            {
                //Stack Trace
                ClsVariables.fnAddStackTrace(pobjException.StackTrace);
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
                        if (DevOpsResult) 
                        { TestContext.Progress.WriteLine($"Test ended with {objLogstatus} – {objErrorMessage}");}
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
                        if (DevOpsResult)
                        {TestContext.Progress.WriteLine($"The status: {objLogstatus} is not supported."); }
                        Console.WriteLine("The status: " + objLogstatus + " is not supported.");
                        break;
                }
            }
            catch (Exception pobjException)
            {
                //Stack Trace
                ClsVariables.fnAddStackTrace(pobjException.StackTrace);
                throw (pobjException);
            }

        }

        /*
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
                        if (DevOpsResult)
                        { TestContext.Progress.WriteLine($"{pstrDescription} - Pass");}
                        objTest.Log(Status.Pass, pstrDescription, MediaEntityBuilder.CreateScreenCaptureFromPath(strSCLocation).Build());
                        break;
                    case "FAIL":
                        TC_Status = false;
                        if (DevOpsResult)
                        { TestContext.Progress.WriteLine($"{pstrDescription} - Fail");}
                        objTest.Log(Status.Fail, pstrDescription, MediaEntityBuilder.CreateScreenCaptureFromPath(strSCLocation).Build());
                        if (pblHardStop)
                            Assert.Fail(pstrHardStopMsg);
                        break;
                    case "INFO":
                        if (DevOpsResult)
                        {TestContext.Progress.WriteLine($"{pstrDescription} - Info"); }
                        objTest.Log(Status.Info, pstrDescription, MediaEntityBuilder.CreateScreenCaptureFromPath(strSCLocation).Build());
                        break;
                    case "WARNING":
                        isWarning = true;
                        if (DevOpsResult)
                        { TestContext.Progress.WriteLine($"{pstrDescription} - Warning");}
                        objTest.Log(Status.Warning, pstrDescription, MediaEntityBuilder.CreateScreenCaptureFromPath(strSCLocation).Build());
                        break;
                }
            }
            else
            {
                switch (pstrStatus.ToUpper())
                {
                    case "PASS":
                        if (DevOpsResult)
                        {TestContext.Progress.WriteLine($"{pstrDescription} - Pass"); }
                        objTest.Log(Status.Pass, pstrDescription);
                        break;
                    case "FAIL":
                        TC_Status = false;
                        if (DevOpsResult)
                        { TestContext.Progress.WriteLine($"{pstrDescription} - Fail");}
                        objTest.Log(Status.Fail, pstrDescription);
                        if (pblHardStop) { Assert.Fail(pstrHardStopMsg); }
                        break;
                    case "INFO":
                        if (DevOpsResult)
                        {TestContext.Progress.WriteLine($"{pstrDescription} - Info"); }
                        objTest.Log(Status.Info, pstrDescription);
                        break;
                    case "WARNING":
                        isWarning = true;
                        if (DevOpsResult)
                        { TestContext.Progress.WriteLine($"{pstrDescription} - Warning");}
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
        */

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

            if (pstrStatus == Status.Fail) 
            { 
                TC_Status = false;
                //Stack Trace
                ClsVariables.fnAddStackTrace($"Step Error: {pstrStepName}, {pstrDescription}");
            }

            if (pstrStatus == Status.Warning) { isWarning = true; }

            if (DevOpsResult)
            {TestContext.Progress.WriteLine($"{pstrStepName}: {pstrDescription}"); }
            objTest.Log(pstrStatus == Status.Skip ? Status.Info : pstrStatus, pstrDescription, ss);
        }

        /// <summary>
        /// Fails the current test scenario by closing the browser and throwing an exception
        /// </summary>
        /// <param name="pstrMessage">Fail message</param>
        public static void fnAssertFail(string pstrMessage)
        {
            TC_Status = false;
            //ClsWebBrowser.fnCloseBrowser();
            throw new Exception($"Assertion Failed: {pstrMessage}");
        }

        /// <summary>
        /// Takes a screenshot during  execution time
        /// </summary>
        /// <returns></returns>
        public static string fnGetScreenshot()
        {
            //string strSCName = "SC_" + ClsDataDriven.strProjectName + "_" + DateTime.Now.ToString("MMddyyyy_hhmmss");
            string strSCName = "SC_" + TestContext.Parameters["GI_ProjectName"].Replace(" ", "_") + "_" + DateTime.Now.ToString("MMddyyyy_hhmmss");

            //To take screenshot
            Screenshot objFile = ((ITakesScreenshot)ClsWebBrowser.objDriver).GetScreenshot();

            //string strFileLocation = ClsDataDriven.strReportLocation + @"\Screenshots\" + strSCName + ".jpg";
            string strFileLocation = BaseReportFolder + @"Screenshots\" + strSCName + ".jpg";
            //To save screenshot
            objFile.SaveAsFile(strFileLocation, ScreenshotImageFormat.Jpeg);

            return strFileLocation;
        }


        /// <summary>
        /// Generates the folder for the report generated
        /// </summary>
        private static void fnFolderSetup()
        {
            try
            {
                string[] strSubFolders = new string[2] { "ScreenShots", TestContext.Parameters["GI_ProjectName"] };
                bool blFExist = System.IO.Directory.Exists(BaseReportFolder);
                if (!blFExist)
                {
                    System.IO.Directory.CreateDirectory(BaseReportFolder);
                }
                else
                    blFExist = false;

                foreach (string strFolder in strSubFolders)
                {
                    blFExist = System.IO.Directory.Exists(BaseReportFolder + strFolder);

                    if (!blFExist)
                    {
                        System.IO.Directory.CreateDirectory(BaseReportFolder + strFolder);
                    }
                }
            }
            catch (Exception e) 
            {
                //Stack Trace
                ClsVariables.fnAddStackTrace(e.StackTrace);
                if (DevOpsResult)
                {TestContext.Progress.WriteLine($"Folder cannot be created, an exception appears: {e.Message}"); }
            }

        }

        public static void fnRenameFolder() 
        {
            try
            {
                var oldPath = BaseReportFolder.Substring(0, BaseReportFolder.Length - 1);
                var newPath = BaseReportFolder.Substring(0, BaseReportFolder.Length - 1) + "_TestPlanSuite_" + TestPlanSuite.Replace("-", "_").Replace(" ", "");
                Directory.Move(@oldPath, @newPath);
                TestContext.Progress.WriteLine($"The rename folder is: {newPath}");
            }
            catch (Exception e)
            {
                //Stack Trace
                ClsVariables.fnAddStackTrace(e.StackTrace);
            }
        }

    }
}
