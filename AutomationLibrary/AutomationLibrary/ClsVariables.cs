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
                    strLocalExecution = Environment.GetEnvironmentVariable("GI_Env_Variable").ToUpper();
                }
                catch (NullReferenceException)
                {
                    //Ignore and continue
                }

                return strLocalExecution.Equals("Local", StringComparison.OrdinalIgnoreCase);
            }
        }

        /// <summary>
        /// Constant to define the Automation Setting driver
        /// </summary>
        public static string strGlobalConfigFile => blLocalExecution ? @"C:\AutomationProjects\AutomationSettings.xlsx" : @"H:\Any\4th_Automation\GlobalIntakeDriver\AutomationSettings.xlsx";

        /// <summary>
        /// Constant static for Edge Driver
        /// </summary>
        public static string strEdgeDriverPath = @"H:\Any\4th_Automation\GlobalIntakeDriver\ExternalLib\";
    }
}
