using AventStack.ExtentReports;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WebDriverManager.DriverConfigs.Impl;

namespace AutomationLibrary
{
    public class ClsWebBrowser
    {
        //private static readonly IDictionary<string, IWebDriver> dicDrivers = new Dictionary<string, IWebDriver>();
        public static IWebDriver _objGlobalDriver = null;
        private static IWebDriver _objDriver;
        private static WebDriverWait _wait;

        /// <summary>
        /// Represents an instance of webdriver
        /// </summary>
        public static IWebDriver objDriver
        {
            get
            {
                return _objDriver;
            }
            private set
            {
                _objDriver = value;
            }
        }

        /// <summary>
        /// Inits an instance of webdriver as Chrome, Edge or Firefox
        /// </summary>
        /// <param name="pstrBrowsername"></param>
        public static void fnInitBrowser(string pstrBrowsername) 
        {
            switch (pstrBrowsername.ToUpper())
            {
                case "CHROME":

                    new WebDriverManager.DriverManager().SetUpDriver(new ChromeConfig());
                    _objDriver = new ChromeDriver();

                    ChromeOptions optionsChrome = new ChromeOptions();
                    optionsChrome.AddArgument("no-sandbox");
                    optionsChrome.AddArgument("start-maximized");

                    //Removing for Driver Manager
                    //_objDriver = new ChromeDriver(ChromeDriverService.CreateDefaultService(), optionsChrome, TimeSpan.FromMinutes(3));

                    _objDriver.Manage().Timeouts().PageLoad.Add(System.TimeSpan.FromSeconds(10));
                    _wait = new WebDriverWait(_objDriver, TimeSpan.FromSeconds(5));
                    _objDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                    _objDriver.Manage().Window.Maximize();
                    _objDriver.Manage().Cookies.DeleteAllCookies();
                    Thread.Sleep(TimeSpan.FromSeconds(5));
                    break;
                case "HEADLESSCHROME":
                    var optionsHeadlessChrome = new ChromeOptions();
                    optionsHeadlessChrome.AddArgument("no-sandbox");
                    optionsHeadlessChrome.AddArgument("window-size=1920,1080");
                    optionsHeadlessChrome.AddArgument("--headless");

                    //Removing for Driver Manager
                    //_objDriver = new ChromeDriver(ChromeDriverService.CreateDefaultService(), optionsHeadlessChrome, TimeSpan.FromMinutes(3));
                    new WebDriverManager.DriverManager().SetUpDriver(new ChromeConfig());
                    _objDriver = new ChromeDriver();

                    //TODO: Next steps can be optimized
                    _objDriver.Manage().Timeouts().PageLoad.Add(System.TimeSpan.FromSeconds(10));
                    _objDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                    _objDriver.Manage().Window.Maximize();
                    _objDriver.Manage().Cookies.DeleteAllCookies();
                    Thread.Sleep(TimeSpan.FromSeconds(5));

                    break;
                case "EDGE":
                    //var strDriverPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                    var strDriverPath = "Provide your path";
                    var optionsEdge = new OpenQA.Selenium.Edge.EdgeOptions();
                    #pragma warning disable CS0618 // Type or member is obsolete
                    optionsEdge.AddAdditionalCapability("UseChromium", true);
                    #pragma warning restore CS0618 // Type or member is obsolete
                    objDriver = new OpenQA.Selenium.Edge.EdgeDriver(OpenQA.Selenium.Edge.EdgeDriverService.CreateDefaultService(strDriverPath), optionsEdge, TimeSpan.FromMinutes(3));
                    objDriver.Manage().Timeouts().PageLoad.Add(System.TimeSpan.FromSeconds(10));
                    _wait = new WebDriverWait(_objDriver, TimeSpan.FromSeconds(5));
                    _objDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                    _objDriver.Manage().Window.Maximize();
                    _objDriver.Manage().Cookies.DeleteAllCookies();
                    Thread.Sleep(TimeSpan.FromSeconds(5));
                    //dicDrivers.Add("Edge", objDriver);
                    break;
                case "FIREFOX":
                    new WebDriverManager.DriverManager().SetUpDriver(new FirefoxConfig());
                    _objDriver = new FirefoxDriver();
                    _objDriver.Manage().Timeouts().PageLoad.Add(System.TimeSpan.FromSeconds(10));
                    _objDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                    _objDriver.Manage().Window.Maximize();
                    break;
            }
        }


        public static IWebDriver fnPrivateSession(string pstrBrowsername, string pstrPreferredLanguaje = "")
        {
            IWebDriver objDriver2 = null;
            switch (pstrBrowsername.ToUpper())
            {
                case "CHROME":
                    ChromeOptions optionsChrome = new ChromeOptions();
                    optionsChrome.AddArgument("no-sandbox");
                    optionsChrome.AddArgument("start-maximized");
                    optionsChrome.AddArgument("incognito");
                    if (pstrPreferredLanguaje != "") { optionsChrome.AddArgument(pstrPreferredLanguaje); }

                    //Remove to add WebDriverManager
                    //objDriver2 = new ChromeDriver(ChromeDriverService.CreateDefaultService(), optionsChrome, TimeSpan.FromMinutes(3));
                    new WebDriverManager.DriverManager().SetUpDriver(new ChromeConfig());
                    objDriver2 = new ChromeDriver();

                    objDriver2.Manage().Timeouts().PageLoad.Add(System.TimeSpan.FromSeconds(10));
                    _wait = new WebDriverWait(objDriver2, TimeSpan.FromSeconds(5));
                    objDriver2.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                    objDriver2.Manage().Window.Maximize();
                    objDriver2.Manage().Cookies.DeleteAllCookies();
                    Thread.Sleep(TimeSpan.FromSeconds(5));
                    break;
            }
            return objDriver2;
        }



        /// <summary>
        /// Navigates to a specific URL
        /// </summary>
        /// <param name="pstrURL"></param>
        public static void fnNavigateToUrl(string pstrURL) 
        {
            objDriver.Url = pstrURL;
        }

        /// <summary>
        /// Closes all the webdriver intances
        /// </summary>
        public static void fnCloseBrowser() 
        {
            ClsReportResult.fnLog("CloseBrowser", "Step - Closing Browser", Status.Info, false);
            _objDriver.Close();
            _objDriver.Quit();
        }


        public static IWebDriver fnChangeDriver(IWebDriver driver) 
        {
            _objGlobalDriver = driver;
            _objDriver = _objGlobalDriver;
            return _objDriver;
        }


    }


}
