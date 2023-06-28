using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationLibrary
{
    public static class ClsVariables
    {
        /// <summary>
        /// Boolean that defines if the current execution is local
        /// </summary>
        public static bool blLocalExecution
        {
            get
            {
                string strLocalExecution = null;
                try
                {
                    //strLocalExecution = Environment.GetEnvironmentVariable("GI_Env_Variable").ToUpper();
                    strLocalExecution = TestContext.Parameters["GI_Env_Variable"].ToUpper().ToString();
                }
                catch (NullReferenceException)
                {
                    //If variable is not defined, will return false by default
                    return false;
                }

                return strLocalExecution.Equals("Local", StringComparison.OrdinalIgnoreCase);
            }
        }

        /// <summary>
        /// Boolean that defines if TC will be skipped
        /// </summary>
        public static bool blSkipExecution
        {
            get
            {
                bool strSkipExecution = false;
                try
                {
                    strSkipExecution = bool.Parse(TestContext.Parameters["GI_Env_SkipTests"]);
                }
                catch (NullReferenceException)
                {
                    return false;
                }

                return strSkipExecution;
            }
        }



        /// <summary>
        /// Constant to define the Automation Setting driver
        /// </summary>
        public static string strGlobalConfigFile => blLocalExecution ? @"C:\AutomationProjects\AutomationSettings.xlsx" : @"H:\Any\4th_Automation\GlobalIntakeDriver\AutomationSettings.xlsx";
        public static string strGlobalRootDirectory => blLocalExecution ? @"C:\AutomationProjects\" : @"H:\Any\4th_Automation\GlobalIntakeDriver\";

        /// <summary>
        /// Constant static for Edge Driver
        /// </summary>
        public static string strEdgeDriverPath = @"H:\Any\4th_Automation\GlobalIntakeDriver\ExternalLib\";

        /// <summary>
        /// Get stacktrace log
        /// </summary>
        public static string strStackTrace = "";

        /// <summary>
        /// 
        /// </summary>
        /// <param name="strStackTrace"></param>
        public static void fnAddStackTrace(string strStackTrace) 
        {
            ClsVariables.strStackTrace = ClsVariables.strStackTrace + "\n\n" + strStackTrace;
        }


    }
}
