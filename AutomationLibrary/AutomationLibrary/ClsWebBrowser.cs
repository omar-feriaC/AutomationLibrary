using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationLibrary
{
    public class ClsWebBrowser
    {
        private static readonly IDictionary<string, IWebDriver> dicDrivers = new Dictionary<string, IWebDriver>();
        private static IWebDriver _objDriver;

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
            if (objDriver == null) 
            {
                switch (pstrBrowsername.ToUpper())
                {
                    case "CHROME":
                        ChromeOptions optionsChrome = new ChromeOptions();
                        optionsChrome.AddArgument("no-sandbox");
                        _objDriver = new ChromeDriver(ChromeDriverService.CreateDefaultService(), optionsChrome, TimeSpan.FromMinutes(3));
                        _objDriver.Manage().Timeouts().PageLoad.Add(System.TimeSpan.FromSeconds(10));
                        _objDriver.Manage().Window.Maximize();
                        dicDrivers.Add("Chrome", objDriver);
                        break;
                    case "EDGE":
                        var strDriverPath = "Provide your path";
                        var optionsEdge = new OpenQA.Selenium.Edge.EdgeOptions();
                        #pragma warning disable CS0618 // Type or member is obsolete
                        optionsEdge.AddAdditionalCapability("UseChromium", true);
                        #pragma warning restore CS0618 // Type or member is obsolete
                        objDriver = new OpenQA.Selenium.Edge.EdgeDriver(OpenQA.Selenium.Edge.EdgeDriverService.CreateDefaultService(strDriverPath), optionsEdge, TimeSpan.FromMinutes(3));
                        objDriver.Manage().Timeouts().PageLoad.Add(System.TimeSpan.FromSeconds(10));
                        objDriver.Manage().Window.Maximize();
                        dicDrivers.Add("Edge", objDriver);
                        break;
                    case "FIREFOX":
                        objDriver = new FirefoxDriver();
                        objDriver.Manage().Timeouts().PageLoad.Add(System.TimeSpan.FromSeconds(10));
                        dicDrivers.Add("FireFox", objDriver);
                        break;
                }
            }
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
        public static void fnCloseAllDrivers() 
        {
            foreach (var key in dicDrivers.Keys) 
            {
                dicDrivers[key].Close();
                dicDrivers[key].Quit();
            }
        }




    }


}
