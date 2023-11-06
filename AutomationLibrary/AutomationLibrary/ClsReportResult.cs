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
using System.Threading;
using System.Threading.Tasks;

namespace AutomationLibrary
{
    public static class ClsReportResult
    {
        
        public static ExtentReports objExtent;
        public static ExtentTest objTest;
        #pragma warning disable CS0618 // Type or member is obsolete
        public static ExtentV3HtmlReporter objHtmlReporter;
        #pragma warning restore CS0618 // Type or member is obsolete
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
                blSuccess = true;

                if (blSuccess)
                {
                    //To create report directory and add HTML report into it
                    string folderID = $"_{Environment.UserName.ToString()}";
                    BaseReportFolder = TestContext.Parameters["GI_ReportLocation"] + DateTime.Now.ToString("MMddyyyy_hhmmss") + folderID + @"\";
                    FullReportFolder = BaseReportFolder + TestContext.Parameters["GI_ProjectName"] + @"\" + TestContext.Parameters["GI_ReportName"] + ".html";
                    fnFolderSetup();

                    //Configure Report
                    #pragma warning disable CS0618 // Type or member is obsolete
                    objHtmlReporter = new ExtentV3HtmlReporter(FullReportFolder);
                    #pragma warning restore CS0618 // Type or member is obsolete
                    objHtmlReporter.Config.ReportName = TestContext.Parameters["GI_ReportName"];
                    objHtmlReporter.Config.DocumentTitle = TestContext.Parameters["GI_ProjectName"] + " - " + TestContext.Parameters["GI_ReportName"];
                    objHtmlReporter.Config.Theme = AventStack.ExtentReports.Reporter.Configuration.Theme.Dark;
                    objHtmlReporter.Config.Encoding = "utf-8";
                    
                    //Adding extra information to Report
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
                string strSCLocation = "";
                try
                {
                    strSCLocation = fnGetScreenshot();
                    ss = MediaEntityBuilder.CreateScreenCaptureFromPath(strSCLocation).Build(); 
                }
                catch 
                {
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    strSCLocation = fnGetScreenshot();
                    ss = MediaEntityBuilder.CreateScreenCaptureFromPath(strSCLocation).Build(); 
                }
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
            throw new Exception($"Assertion Failed: {pstrMessage}");
        }

        /// <summary>
        /// Takes a screenshot during  execution time
        /// </summary>
        /// <returns></returns>
        public static string fnGetScreenshot()
        {
            string strFileLocation = "";
            try
            {
                string strSCName = "SC_" + TestContext.Parameters["GI_ProjectName"].Replace(" ", "_") + "_" + DateTime.Now.ToString("MMddyyyy_hhmmss");
                //To take screenshot
                Screenshot objFile = ((ITakesScreenshot)ClsWebBrowser.objDriver).GetScreenshot();
                strFileLocation = BaseReportFolder + @"Screenshots\" + strSCName + ".jpg";
                //To save screenshot
                objFile.SaveAsFile(strFileLocation, ScreenshotImageFormat.Jpeg);
            }
            catch 
            { }
            
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

        /// <summary>
        /// Renames the report path, but screenshots are loss when this action is executed.
        /// </summary>
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
