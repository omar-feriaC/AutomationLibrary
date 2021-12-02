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
        /// Constant to define the Automation Setting driver
        /// </summary>
        public static string strGlobalConfigFile
        {
            get
            {
                string blLocalExecution = null;
                try
                {
                    blLocalExecution = Environment.GetEnvironmentVariable("GI_Env_Variable");

                }
                catch (NullReferenceException)
                {
                    //Ignore and continue
                }

                string path;
                switch (blLocalExecution)
                {
                    case "LOCAL":
                        path = @"C:\AutomationProjects\AutomationSettings.xlsx";
                        break;
                    default:
                        path = @"H:\Any\4th_Automation\GlobalIntakeDriver\AutomationSettings.xlsx";
                        break;
                }
                return path;
            }
        }

        /// <summary>
        /// Constant static for Edge Driver
        /// </summary>
        public static string strEdgeDriverPath = @"H:\Any\4th_Automation\GlobalIntakeDriver\ExternalLib\";
    }
}
