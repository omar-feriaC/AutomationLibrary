using AventStack.ExtentReports;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AutomationLibrary
{
    public class ClsWebElements
    {
        private static DefaultWait<IWebDriver> objFluentWait;
        private static WebDriverWait objExplicitWait;
        private static string strAction = "";

        /// <summary>
        /// Returns a list of web elements
        /// </summary>
        /// <param name="by"></param>
        /// <returns></returns>
        public static IList<IWebElement> fnGetWebList(By by)
        {
            try
            {
                IList<IWebElement> objElement = ClsWebBrowser.objDriver.FindElements(by);
                return objElement;
            }
            catch (Exception pobjException) 
            {
                fnExceptionHandling(pobjException, "Ilist<WebElement>: doesn't exist in the page.", true);
                return null;
            }
        }

        /// <summary>
        /// Returns a list of web elements
        /// </summary>
        /// <param name="pstrLocator"></param>
        /// <returns></returns>
        public static IList<IWebElement> fnGetWebList(string pstrLocator)
        {
            try
            {
                IList<IWebElement> objElement = ClsWebBrowser.objDriver.FindElements(By.XPath(pstrLocator));
                return objElement;
            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException, "Ilist<WebElement>: doesn't exist in the page", true);
                return null;
            }
        }

        /// <summary>
        /// Returns a web element of the page
        /// </summary>
        /// <param name="by"></param>
        /// <returns></returns>
        public static IWebElement fnGetWebElement(By by)
        {
            try
            {
                IWebElement pobjElement = ClsWebBrowser.objDriver.FindElement(by);
                return pobjElement;
            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException, "WebElement: doesn't exist in the page.", true);
                return null;
            }
        }

        /// <summary>
        /// Returns a web element of the page
        /// </summary>
        /// <param name="pstrLocator"></param>
        /// <returns></returns>
        public static IWebElement fnGetWebElement(string pstrLocator) => fnGetWebElement(By.XPath(pstrLocator));


        /// <summary>
        /// Get Web Element Given a parent and a relative locator
        /// </summary>
        /// <param name="element">The parent element</param>
        /// <param name="locator">The locator</param>
        /// <param name="descr">Optional description of the element to find</param>
        /// <returns>The WebElement</returns>
        public static IWebElement fnGetWebElement(IWebElement element, By locator, string descr = "")
        {
            try
            {
                IWebElement pobjElement = element.FindElement(locator);
                return pobjElement;
            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException, $"WebElement not found{(string.IsNullOrEmpty(descr) ? "" : $": {descr}")}", true);
                return null;
            }
        }

        /// <summary>
        /// Get Web Element Given a parent and a relative XPath
        /// </summary>
        /// <param name="element">The parent element</param>
        /// <param name="xpath">The XPath locator</param>
        /// <param name="descr">Optional description of the element to find</param>
        /// <returns>The WebElement</returns>
        public static IWebElement fnGetWebElement(IWebElement element, string xpath, string descr = "") => fnGetWebElement(element, xpath.StartsWith(".") ? By.XPath(xpath) : By.XPath("." + xpath), descr);

        /// <summary>
        /// Executes an action specified
        /// </summary>
        /// <param name="pobjWebElement"></param>
        /// <param name="pstrAction"></param>
        /// <param name="pstrTextEnter"></param>
        /// <param name="Timeout"></param>
        /// <returns></returns>
        public static object fnGetFluentWait(IWebElement pobjWebElement, string pstrAction, string pstrTextEnter = "", int Timeout = 5)
        {
            if (pobjWebElement == null) throw new NullReferenceException("Please make sure to send a not null WebElement instance before calling this function");

            objFluentWait = new DefaultWait<IWebDriver>(ClsWebBrowser.objDriver);
            objFluentWait.Timeout = TimeSpan.FromSeconds(Timeout);
            objFluentWait.PollingInterval = TimeSpan.FromMilliseconds(250);
            objFluentWait.IgnoreExceptionTypes(typeof(WebDriverTimeoutException), typeof(SuccessException));

            switch (pstrAction)
            {
                case "Displayed":
                    objFluentWait.Until(x => pobjWebElement.Displayed);
                    break;
                case "Click":
                    objFluentWait.Until(x => pobjWebElement).Click();
                    break;
                case "SendKeys":
                    objFluentWait.Until(x => pobjWebElement).SendKeys(pstrTextEnter);
                    break;
                case "Clear":
                    objFluentWait.Until(x => pobjWebElement).Clear();
                    break;
                case "CustomSendKeys":
                    //fnScrollToV2(pobjWebElement, "Scroll to element", false);
                    Actions action = new Actions(ClsWebBrowser.objDriver);
                    pobjWebElement.Click();
                    action.KeyDown(Keys.Control).SendKeys(Keys.Home).Perform();
                    pobjWebElement.Clear();
                    Thread.Sleep(TimeSpan.FromMilliseconds(500));
                    pobjWebElement.SendKeys(Keys.Delete);
                    pobjWebElement.SendKeys(pstrTextEnter);
                    //Thread.Sleep(TimeSpan.FromMilliseconds(500));
                    break;
            }
            return objFluentWait;
        }

        /// <summary>
        /// Verify if an element exist with and explicit wait
        /// </summary>
        /// <param name="pstrLocator"></param>
        /// <param name="pstrOption"></param>
        /// <param name="Timeout"></param>
        /// <returns></returns>
        public static bool fnGetExplicitWait(By by, string pstrOption, int Timeout = 5)
        {
            bool pblStatus = false;
            IWebElement objWebElement;
            try
            {
                objExplicitWait = new WebDriverWait(ClsWebBrowser.objDriver, TimeSpan.FromSeconds(Timeout));
                objExplicitWait.IgnoreExceptionTypes(typeof(WebDriverTimeoutException));

                switch (pstrOption)
                {
                    case "ElementExists":
                        objWebElement = objExplicitWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(by));
                        pblStatus = true;
                        break;
                    case "ElementVisible":
                        objWebElement = objExplicitWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(by));
                        pblStatus = true;
                        break;
                }
                return pblStatus;

            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException);
                return pblStatus;
            }
        }

        /// <summary>
        /// Verify if an element exist with and explicit wait
        /// </summary>
        /// <param name="pstrLocator"></param>
        /// <param name="pstrOption"></param>
        /// <param name="Timeout"></param>
        /// <returns></returns>
        public static bool fnGetExplicitWait(string pstrLocator, string pstrOption, int Timeout = 5)
        {
            bool pblStatus = false;
            IWebElement objWebElement;
            try
            {
                objExplicitWait = new WebDriverWait(ClsWebBrowser.objDriver, TimeSpan.FromSeconds(Timeout));
                objExplicitWait.IgnoreExceptionTypes(typeof(WebDriverTimeoutException));

                switch (pstrOption)
                {
                    case "ElementExists":
                        objWebElement = objExplicitWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath(pstrLocator)));
                        pblStatus = true;
                        break;
                    case "ElementVisible":
                        objWebElement = objExplicitWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(pstrLocator)));
                        pblStatus = true;
                        break;
                }
                return pblStatus;

            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException);
                return pblStatus;
            }
        }

        /// <summary>
        /// Wait to page to ne loaded
        /// </summary>
        /// <param name="pobjWebElement"></param>
        /// <param name="pstrPage"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        /// <returns></returns>
        public static bool fnPageLoad(IWebElement pobjWebElement, string pstrPage, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "PageLoad Failed and HardStop defined")
        {
            if (pobjWebElement == null) throw new NullReferenceException("Please make sure to send a not null WebElement instance before calling this function");

            //ClsReportResult clsRR = new clsReportResult();
            bool blResult = false;
            try
            {
                ClsReportResult.fnLog("PageLoad", "Step - PageLoad in Page: " + pstrPage, Status.Info, false);
                strAction = "Displayed";
                IJavaScriptExecutor objJS = (IJavaScriptExecutor)ClsWebBrowser.objDriver;
                IWait<IWebDriver> wait = new WebDriverWait(ClsWebBrowser.objDriver, TimeSpan.FromSeconds(10));
                wait.Until(wd => objJS.ExecuteScript("return document.readyState").ToString() == "complete");
                fnGetFluentWait(pobjWebElement, strAction);
                ClsReportResult.fnLog("PageLoadPass", "The Page is loaded for the Page: " + pstrPage, Status.Pass, pblScreenShot);
                blResult = true;
            }
            catch (Exception pobjException)
            {
                ClsReportResult.fnLog("PageLoadFail", "The Page is not loaded for the Page: " + pstrPage, Status.Fail, true);
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        public static bool fnPageLoad(By by, string pstrPage, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "PageLoad Failed and HardStop defined")
        {
            //if (pobjWebElement == null) throw new NullReferenceException("Please make sure to send a not null WebElement instance before calling this function");

            //ClsReportResult clsRR = new clsReportResult();
            bool blResult = false;
            try
            {
                ClsReportResult.fnLog("PageLoad", "Step - PageLoad in Page: " + pstrPage, Status.Info, false);
                strAction = "Displayed";
                IJavaScriptExecutor objJS = (IJavaScriptExecutor)ClsWebBrowser.objDriver;
                IWait<IWebDriver> wait = new WebDriverWait(ClsWebBrowser.objDriver, TimeSpan.FromSeconds(10));
                wait.Until(wd => objJS.ExecuteScript("return document.readyState").ToString() == "complete");
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(by));
                fnGetFluentWait(ClsWebElements.fnGetWebElement(by), strAction);
                ClsReportResult.fnLog("PageLoadPass", "The Page is loaded for the Page: " + pstrPage, Status.Pass, pblScreenShot);
                blResult = true;
            }
            catch (Exception pobjException)
            {
                ClsReportResult.fnLog("PageLoadFail", "The Page is not loaded for the Page: " + pstrPage, Status.Fail, true);
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        /// <summary>
        /// Send a text to input fields
        /// </summary>
        /// <param name="pobjWebElement"></param>
        /// <param name="pstrField"></param>
        /// <param name="pstrTextEnter"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        /// <returns></returns>
        public static bool fnSendKeys(IWebElement pobjWebElement, string pstrField, string pstrTextEnter, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "SendKeys Failed and HardStop defined")
        {
            //ClsReportResult clsRR = new clsReportResult();
            bool blResult = false;

            try
            {
                ClsReportResult.fnLog("SendKeys", "Step - Sendkeys: " + pstrTextEnter + " to field: " + pstrField, Status.Info, false);
                strAction = "SendKeys";
                fnGetFluentWait(pobjWebElement, strAction, pstrTextEnter);
                ClsReportResult.fnLog("SendKeysPass", "The SendKeys for: " + pstrField + " with value: " + pstrTextEnter + " was done successfully.", Status.Pass, pblScreenShot);
                blResult = true;
            }
            catch (Exception pobjException)
            {
                ClsReportResult.fnLog("SendKeysFail", "The SendKeys for: " + pstrField + " with value: " + pstrTextEnter + " has failed.", Status.Fail, true);
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        /// <summary>
        /// Sends a text forcing the action
        /// </summary>
        /// <param name="pobjWebElement"></param>
        /// <param name="pstrField"></param>
        /// <param name="pstrTextEnter"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        /// <returns></returns>
        public static bool fnCustomSendKeys(IWebElement pobjWebElement, string pstrField, string pstrTextEnter, bool pblScreenShot = true)
        {
            //ClsReportResult clsRR = new clsReportResult();
            bool blResult = false;

            try
            {
                ClsReportResult.fnLog("SendKeys", "Step - Sendkeys: " + pstrTextEnter + " to field: " + pstrField, Status.Info, false);
                strAction = "CustomSendKeys";
                fnGetFluentWait(pobjWebElement, strAction, pstrTextEnter);
                ClsReportResult.fnLog("SendKeysPass", "The SendKeys for: " + pstrField + " with value: " + pstrTextEnter + " was done successfully.", Status.Pass, pblScreenShot);
                blResult = true;
            }
            catch (Exception pobjException)
            {
                ClsReportResult.fnLog("SendKeysFail", "The SendKeys for: " + pstrField + " with value: " + pstrTextEnter + " has failed.", Status.Fail, true);
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        /// <summary>
        /// Click to web element
        /// </summary>
        /// <param name="pobjWebElement"></param>
        /// <param name="pstrElement"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        /// <returns></returns>
        public static bool fnClick(IWebElement pobjWebElement, string pstrElement, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Click Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                ClsReportResult.fnLog("Click", "Step - Click on " + pstrElement, Status.Info, false);
                strAction = "Click";
                fnGetFluentWait(pobjWebElement, strAction);
                ClsReportResult.fnLog("ClickPass", "Click on " + pstrElement + " was done successfully.", Status.Pass, pblScreenShot);
                blResult = true;
            }
            catch (Exception pobjException)
            {
                ClsReportResult.fnLog("ClickFail", "The click to the element is not working for: " + pstrElement, Status.Fail, true);
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        /// <summary>
        /// Double click to and web element
        /// </summary>
        /// <param name="pobjWebElement"></param>
        /// <param name="pstrElement"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        /// <returns></returns>
        public static bool fnDoubleClick(IWebElement pobjWebElement, string pstrElement, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "DoubleClick Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                ClsReportResult.fnLog("DoubleClick", "Step - Double Clic on " + pstrElement, Status.Info, false);
                strAction = "Displayed";
                fnGetFluentWait(pobjWebElement, strAction);
                Actions actions = new Actions(ClsWebBrowser.objDriver);
                actions.DoubleClick(pobjWebElement).Perform();
                ClsReportResult.fnLog("DoubleClickPass", "Double Click on " + pstrElement + " was done successfully.", Status.Pass, pblScreenShot);
                blResult = true;
            }
            catch (Exception pobjException)
            {
                ClsReportResult.fnLog("DoubleClickPass", "Couldn't Double Click on " + pstrElement, Status.Fail, true);
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        /// <summary>
        /// Get the attribute of an element
        /// </summary>
        /// <param name="pobjWebElement"></param>
        /// <param name="pstrElement"></param>
        /// <param name="pstrAttName"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        /// <returns></returns>
        public static string fnGetAttribute(IWebElement pobjWebElement, string pstrElement, string pstrAttName, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "GetAttribute Failed and HardStop defined")
        {
            string strAttributeContent = "";
            try
            {
                ClsReportResult.fnLog("GetAttribute", "Step - Get Attribue " + pstrAttName + " from " + pstrElement, Status.Info, false);
                strAttributeContent = pobjWebElement.GetAttribute(pstrAttName);
                ClsReportResult.fnLog("GetAttributePass", "Get Attribute " + pstrAttName + " from Element: " + pstrElement + "  was done successfully.", Status.Pass, pblScreenShot);
            }
            catch (Exception pobjException)
            {
                ClsReportResult.fnLog("GetAttributeFail", "Get Attribute " + pstrAttName + " from Element: " + pstrElement + " was not done.", Status.Fail, false);
                fnExceptionHandling(pobjException);
            }
            return strAttributeContent;
        }

        /// <summary>
        /// Clear the text of a web element
        /// </summary>
        /// <param name="pobjWebElement"></param>
        /// <param name="pstrField"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        /// <returns></returns>
        public static bool fnClear(IWebElement pobjWebElement, string pstrField, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Clear Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                ClsReportResult.fnLog("Clear", "Step - Clear to field: " + pstrField, Status.Info, false);
                strAction = "Clear";
                fnGetFluentWait(pobjWebElement, strAction);
                Thread.Sleep(TimeSpan.FromSeconds(2));
                string testTXT = pobjWebElement.Text;
                if (pobjWebElement.Text.Equals("") || pobjWebElement.Text.Equals(null))
                {
                    ClsReportResult.fnLog("ClearPass", "Clear to field" + pstrField + " was done successfully.", Status.Pass, pblScreenShot);
                    blResult = true;
                }
                else
                    throw new ArgumentException("ClearFail");
            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        /// <summary>
        /// Verify if a string contains a sub string
        /// </summary>
        /// <param name="pstrStepName"></param>
        /// <param name="pstrParentString"></param>
        /// <param name="pstrSubString"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        /// <returns></returns>
        public static bool fnContainsText(string pstrStepName, string pstrParentString, string pstrSubString, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Contains Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                ClsReportResult.fnLog("Contains", "Step - Contains: " + pstrSubString + " in the string: " + pstrParentString, Status.Info, false);
                if (pstrParentString.Contains(pstrSubString))
                {
                    ClsReportResult.fnLog("ContainsTextPass", pstrStepName, Status.Pass, pblScreenShot);
                    blResult = true;
                }
                else
                    throw new ArgumentException("ContainsTextFail");
            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        /// <summary>
        /// Verifies if a web element exist
        /// </summary>
        /// <param name="pstrStepName"></param>
        /// <param name="pstrLocator"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        /// <returns></returns>
        public static bool fnElementExist(string pstrStepName, string pstrLocator, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Element Exist Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                bool blElementExist = fnGetExplicitWait(pstrLocator, "ElementExists");

                if (blElementExist)
                {
                    ClsReportResult.fnLog("ElementExistPass", "The element " + pstrStepName + " exist in the page", Status.Pass, pblScreenShot);
                    blResult = true;
                }
                else
                    throw new ArgumentException("ElementExistFail");
            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        /// <summary>
        /// Verifies if a web element exist
        /// </summary>
        /// <param name="pstrStepName"></param>
        /// <param name="by"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        /// <returns></returns>
        public static bool fnElementExist(string pstrStepName, By by, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Element Exist Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                bool blElementExist = fnGetExplicitWait(by, "ElementExists");

                if (blElementExist)
                {
                    ClsReportResult.fnLog("ElementExistPass", "The element " + pstrStepName + " exist in the page", Status.Pass, pblScreenShot);
                    blResult = true;
                }
                else
                    throw new ArgumentException("ElementExistFail");
            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        /// <summary>
        /// Verifies if a web element does not exist
        /// </summary>
        /// <param name="pstrStepName"></param>
        /// <param name="pstrLocator"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        /// <returns></returns>
        public static bool fnElementNotExist(string pstrStepName, string pstrLocator, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Element Not Exist Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                bool blElementExist = fnGetExplicitWait(pstrLocator, "ElementExists");

                if (!blElementExist)
                {
                    ClsReportResult.fnLog("ElementNotExistPass", "Element " + pstrStepName + " not exist in the page", Status.Pass, pblScreenShot);
                    blResult = true;
                }
                else
                    throw new ArgumentException("ElementNotExistFail");

            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        /// <summary>
        /// Verifies if a web element does not exist
        /// </summary>
        /// <param name="pstrStepName"></param>
        /// <param name="by"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        /// <returns></returns>
        public static bool fnElementNotExist(string pstrStepName, By by, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Element Not Exist Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                bool blElementExist = fnGetExplicitWait(by, "ElementExists");

                if (!blElementExist)
                {
                    ClsReportResult.fnLog("ElementNotExistPass", "Element " + pstrStepName + " not exist in the page", Status.Pass, pblScreenShot);
                    blResult = true;
                }
                else
                    throw new ArgumentException("ElementNotExistFail");
            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        /// <summary>
        /// Verifies that two strings are equals
        /// </summary>
        /// <param name="pstrStepName"></param>
        /// <param name="pstrExpectedString"></param>
        /// <param name="pstrActualString"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        /// <returns></returns>
        public static bool fnVerifyText(string pstrStepName, string pstrExpectedString, string pstrActualString, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Verify Text Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                ClsReportResult.fnLog("VerifyText", "Step - Verify Text, Expected: " + pstrExpectedString + ", Actual: " + pstrActualString, Status.Info, false);
                if (pstrExpectedString.Equals(pstrActualString))
                {
                    ClsReportResult.fnLog("VerifyTextPass", pstrStepName, Status.Pass, pblScreenShot);
                    blResult = true;
                }
                else
                    throw new ArgumentException("VerifyTextFail");
            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        /// <summary>
        /// Select an element of a lis
        /// </summary>
        /// <param name="pstrStepName"></param>
        /// <param name="pobjWebElement"></param>
        /// <param name="pstrValue"></param>
        /// <param name="pstrValueType"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        /// <returns></returns>
        public static bool fnSelectList(string pstrStepName, IWebElement pobjWebElement, string pstrValue, string pstrValueType = "value", bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Select List Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                ClsReportResult.fnLog("SelectDropdown", "Step - Select Dropdown: " + pstrStepName + " With Value: " + pstrValue, Status.Info, false);
                var selectElement = new SelectElement(pobjWebElement);
                switch (pstrValueType.ToUpper())
                {
                    case "VALUE":
                        selectElement.SelectByValue(pstrValue);
                        break;
                    case "TEXT":
                        selectElement.SelectByText(pstrValue);
                        break;
                    case "INDEX":
                        selectElement.SelectByIndex(Int32.Parse(pstrValue));
                        break;
                }
                ClsReportResult.fnLog("SelectListPass", pstrStepName, Status.Pass, pblScreenShot);
                blResult = true;

            }
            catch (Exception pobjException)
            {
                ClsReportResult.fnLog("SelectListFail", pstrStepName, Status.Fail, true);
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        /// <summary>
        /// Select an element of a select
        /// </summary>
        /// <param name="pstrStepName"></param>
        /// <param name="pobjWebElement"></param>
        /// <param name="pstrValue"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        /// <returns></returns>
        public static bool fnVerifySelectedItem(string pstrStepName, IWebElement pobjWebElement, string pstrValue, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Selected List Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                ClsReportResult.fnLog("VerifySelectedDropdown", "Step - Selected Dropdown: " + pstrStepName + " With Value: " + pstrValue, Status.Info, false);
                string pstrSelectItem = new SelectElement(pobjWebElement).SelectedOption.GetAttribute("value");
                if (pstrValue == pstrSelectItem)
                {
                    ClsReportResult.fnLog("SelectedItemPass", pstrStepName, Status.Pass, pblScreenShot);
                    blResult = true;
                }
                else
                    throw new ArgumentException("VerifySelectedItemFail");
            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException, pstrStepName);
            }
            return blResult;
        }

        /// <summary>
        /// Select a checkbox element
        /// <summary>
        /// <param name="pstrStepName"></param>
        /// <param name="pobjWebElement"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        /// <returns></returns>
        public static bool fnSelectCheckBox(string pstrStepName, IWebElement pobjWebElement, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Select CheckBox Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                ClsReportResult.fnLog("SelectCheckBox", "Step - " + pstrStepName, Status.Info, false);
                strAction = "Click";
                fnGetFluentWait(pobjWebElement, strAction);
                Thread.Sleep(TimeSpan.FromSeconds(3));
                ClsReportResult.fnLog("SelectCheckBoxPass", pstrStepName, Status.Pass, pblScreenShot);
                blResult = true;
            }
            catch (Exception pobjException)
            {
                ClsReportResult.fnLog("SelectCheckBoxFail", pstrStepName, Status.Fail, true);
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        /// <summary>
        /// Selects a radio button
        /// </summary>
        /// <param name="pstrStepName"></param>
        /// <param name="pobjWebElement"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        /// <returns></returns>
        public static bool fnSelectRadioBtn(string pstrStepName, IWebElement pobjWebElement, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Select Radio Button Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                ClsReportResult.fnLog("SelectRadioBtn", "Step - " + pstrStepName, Status.Info, false);
                strAction = "Click";
                fnGetFluentWait(pobjWebElement, strAction);
                Thread.Sleep(TimeSpan.FromSeconds(3));
                if (pobjWebElement.Selected)
                {
                    ClsReportResult.fnLog("SelectRadioBtnPass", pstrStepName, Status.Pass, pblScreenShot);
                    blResult = true;
                }
                else
                    throw new ArgumentException("SelectRadioBtnFail");
            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException, pstrStepName);
            }
            return blResult;
        }

        /// <summary>
        /// Verify a list of web elements based on tag name
        /// </summary>
        /// <param name="pstrStepName"></param>
        /// <param name="pobjWebElement"></param>
        /// <param name="pstrListValues"></param>
        /// <param name="pstrListType"></param>
        /// <param name="pblScreenShot"></param>
        /// <param name="pblHardStop"></param>
        /// <param name="pstrHardStopMsg"></param>
        /// <returns></returns>
        public static bool fnVerifyList(string pstrStepName, IWebElement pobjWebElement, string[] pstrListValues, string pstrListType = "option", bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Verify List Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                bool blElementsFound = true;
                ClsReportResult.fnLog("VerifyList", "Step - " + pstrStepName, Status.Info, false);
                IList<IWebElement> objListWE = pobjWebElement.FindElements(By.TagName(pstrListType));
                for (int i = 0; i < pstrListValues.Count(); i++)
                {
                    bool blElementFound = false;
                    foreach (IWebElement objWE in objListWE)
                    {
                        if (pstrListValues[i] == objWE.Text)
                        {
                            blElementFound = true;
                        }
                    }

                    if (blElementFound)
                    {
                        ClsReportResult.fnLog("VerifyListItems", "Element from the List: " + pstrListValues[i] + " is Displayed", Status.Info, false);
                    }
                    else
                    {
                        ClsReportResult.fnLog("VerifyListItems", "Element from the List: " + pstrListValues[i] + " is not Displayed", Status.Info, false);
                        blElementsFound = false;
                    }
                }

                if (blElementsFound)
                {
                    ClsReportResult.fnLog("VerifyListPass", "All Elements from the List are Displayed", Status.Pass, pblScreenShot);
                    blResult = true;
                }
                else
                    throw new ArgumentException("VerifyListFail");
            }
            catch (Exception pobjException)
            {
                ClsReportResult.fnLog("VerifyListPass", "Some Elements from the List are not Displayed", Status.Fail, pblScreenShot);
                fnExceptionHandling(pobjException, pstrStepName);
            }
            return blResult;
        }

        /// <summary>
        /// Created to scroll in pages as needed and make a specific element visible
        /// </summary>
        /// <param name="driver">The WebDriver</param>
        /// <param name="element">The elelemt to scroll to</param>
        public static void fnScrollToV2(IWebElement pobjWebElement, string pstrField, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Scroll To Failed and HardStop defined")
        {
            try
            {
                ClsReportResult.fnLog("ScrollTo", "Step - Scroll to element: " + pstrField, Status.Info, false);
                Thread.Sleep(TimeSpan.FromSeconds(2));
                new Actions(ClsWebBrowser.objDriver)
                    .MoveToElement(pobjWebElement)
                    .Build()
                    .Perform();
                ClsReportResult.fnLog("ScrollToPass", "Scrolled to element: " + pstrField, Status.Pass, pblScreenShot);
            }
            catch (Exception pobjException)
            {
                ClsReportResult.fnLog("ScrollToFailed", "Failed Scroll to element: " + pstrField, Status.Fail, true);
                fnExceptionHandling(pobjException);
            }
        }

        /// <summary>
        /// Created to scroll in pages as needed and make a specific element visible, Javascript version
        /// </summary>
        /// <param name="driver">The WebDriver</param>
        /// <param name="element">The elelemt to scroll to</param>
        public static void fnJsScrollTo(IWebElement pobjWebElement, string pstrField, bool pblScreenShot = true, bool alignToTop = false)
        {
            //ClsReportResult clsRR = new clsReportResult();
            try
            {
                ClsReportResult.fnLog("ScrollTo", "Step - Scroll to element: " + pstrField, Status.Info, false);
                ((IJavaScriptExecutor)ClsWebBrowser.objDriver).ExecuteScript($"arguments[0].scrollIntoView({alignToTop.ToString()});", pobjWebElement);
                ClsReportResult.fnLog("ScrollToPass", "Scrolled to element: " + pstrField, Status.Pass, pblScreenShot);
            }
            catch (Exception pobjException)
            {
                ClsReportResult.fnLog("ScrollToFailed", "Failed Scroll to element: " + pstrField, Status.Fail, true);
                fnExceptionHandling(pobjException);
            }
        }

        /// <summary>
        /// Created to scroll in pages as needed and make a specific element visible in the middle of the screen, Javascript version
        /// </summary>
        /// <param name="driver">The WebDriver</param>
        /// <param name="element">The elelemt to scroll to</param>
        public static void fnJsScrollToIntoMiddle(IWebElement pobjWebElement, string pstrField, bool pblScreenShot = true)
        {
            try
            {
                ClsReportResult.fnLog("ScrollTo", "Step - Scroll to element: " + pstrField, Status.Info, false);
                var scrollElementIntoMiddle = "var viewPortHeight = Math.max(document.documentElement.clientHeight, window.innerHeight || 0); var elementTop = arguments[0].getBoundingClientRect().top; window.scrollBy(0, elementTop-(viewPortHeight/2));";
                ((IJavaScriptExecutor)ClsWebBrowser.objDriver).ExecuteScript(scrollElementIntoMiddle, pobjWebElement);
                ClsReportResult.fnLog("ScrollToPass", "Scrolled to element: " + pstrField, Status.Pass, pblScreenShot);
            }
            catch (Exception pobjException)
            {
                ClsReportResult.fnLog("ScrollToFailed", "Failed Scroll to element: " + pstrField, Status.Fail, true);
                fnExceptionHandling(pobjException);
            }
        }

        /// <summary>
        /// Wait to the element to be clickable
        /// </summary>
        /// <param name="by"></param>
        /// <param name="Timeout"></param>
        /// <returns></returns>
        public static bool fnWaitToBeClickable(By by, int Timeout = 5)
        {
            bool pblStatus = false;
            IWebElement objWebElement;
            try
            {
                objExplicitWait = new WebDriverWait(ClsWebBrowser.objDriver, TimeSpan.FromSeconds(Timeout));
                objExplicitWait.IgnoreExceptionTypes(typeof(WebDriverTimeoutException));
                objWebElement = objExplicitWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(by));
                pblStatus = true;
                return pblStatus;

            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException);
                return pblStatus;
            }
        }

        public static void fnExceptionHandling(Exception pobjException, string pstrStepName = "", bool pblHardStop = false, string pstrHardStopMsg = "Failed Step and HardStop defined")
        {
            //ClsReportResult clsRR = new clsReportResult();
            switch (pobjException.Message.ToString())
            {
                case "SendKeysFail":
                    ClsReportResult.fnLog("SendKeysFail", "SendKeys action Fail", Status.Fail, true);
                    break;
                case "ClearFail":
                    ClsReportResult.fnLog("ClearFail", "Clear action Fail", Status.Fail, true);
                    break;
                case "ElementExistFail":
                    ClsReportResult.fnLog("ElementExistFail", "Element exist verification failed", Status.Fail, true);
                    break;
                case "ElementNotExistFail":
                    ClsReportResult.fnLog("ElementNotExistFail", "Element not exist verification failed", Status.Fail, true);
                    break;
                case "ContainsTextFail":
                    ClsReportResult.fnLog("ContainsTextFail", "Contains text verification failed", Status.Fail, true);
                    break;
                case "VerifyTextFail":
                    ClsReportResult.fnLog("VerifyTextFail", "Verify text verification failed", Status.Fail, true);
                    break;
                case "VerifySelectedItemFail":
                    ClsReportResult.fnLog("SelectedItemFail", "Coverage selected verification failed", Status.Fail, true);
                    break;
                case "SelectRadioBtnFail":
                    ClsReportResult.fnLog("SelectRadioBtnFail", "Select Radio Button verification failed", Status.Fail, true);
                    break;
                case "Timed out after 10 seconds":
                    if (pobjException.InnerException.ToString().Contains("no such element: Unable to locate element"))
                    {
                        ClsReportResult.fnLog("NoSuchElement", "WebElement doesn't exist or is incorrect", Status.Info, false);
                    }
                    break;
                default:
                    ClsReportResult.fnLog("Exception", $"{pstrStepName}, Exception => Message({pobjException.Message.ToString()}), Stack Trace({pobjException.StackTrace.ToString()})", Status.Fail, true);
                    break;
            }
        }

    }

}
