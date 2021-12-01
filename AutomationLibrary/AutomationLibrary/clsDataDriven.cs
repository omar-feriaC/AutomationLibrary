using Microsoft.Office.Interop.Excel;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationLibrary
{
    public class ClsDataDriven
    {

        public Application objExcel;
        public Workbook objWkb;
        public Worksheet objSheet;
        public bool blExecutionFlag = true;

        public static string strDataDriverLocation;
        public static string strProjectName;
        public static string strReportLocation;
        public static string strDriverSheet;
        public static string strReportName;
        public static string strReportEnv;
        public static string strBrowser;
        public static string strExecutionDate;
        public static string strExecutionSet;
        public static int intCountTests;
        public static string[,] objTestCases;
        public static string[,] objContent;

        public static int intCountX;
        public static int intCountY;



        public static string strConfigFile = ClsVariables.strGlobalConfigFile;

        public void fnOpenExcel(bool blRead = true)
        {
            objExcel = new Application();
            objWkb = objExcel.Workbooks.Open(strConfigFile, false, blRead,
                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                            Type.Missing, false, Type.Missing, Type.Missing, Type.Missing,
                                            Type.Missing, Type.Missing);
        }

        public void fnCloseExcel()
        {
            objExcel.DisplayAlerts = false;
            objWkb.Save();
            objWkb.Close();
            objExcel.Quit();
        }

        public bool fnAutomationSettings()
        {
            //try
            //{
            //    fnExcelMethod();
            //}
            //catch (Exception pobjException) 
            //{
            //    fnSSLightMethod();
            //}

            fnSSLightMethod();

            fnFolderSetup();

            return blExecutionFlag;
        }

        public void fnExcelMethod()
        {

            fnOpenExcel();
            objSheet = objWkb.Worksheets["Configuration"] as Worksheet;

            strProjectName = objSheet.Cells[2, 1].Text();
            strReportLocation = objSheet.Cells[2, 2].Text() + DateTime.Now.ToString("MMddyyyy_hhmmss");
            strDataDriverLocation = objSheet.Cells[2, 3].Text();
            strDriverSheet = objSheet.Cells[2, 4].Text();
            string strExecutionRow;

            objSheet = objWkb.Worksheets["Executions"] as Worksheet;

            int x = 2;

            do
            {
                strExecutionRow = objSheet.Cells[x, 5].Text();

                if (strExecutionRow.ToUpper() == "NEW")
                {
                    //strExecutionDate = objSheet.Cells[x, 1].Text();
                    strReportName = objSheet.Cells[x, 2].Text();
                    strReportEnv = objSheet.Cells[x, 3].Text();
                    strExecutionSet = objSheet.Cells[x, 4].Text();
                    strBrowser = objSheet.Cells[x, 6].Text();

                    blExecutionFlag = true;
                    break;
                }
                else
                    blExecutionFlag = false;
                x++;
            } while (objSheet.Cells[x, 1].Text() != "");

            objSheet = objWkb.Worksheets["TestCases"] as Worksheet;

            x = 2;
            intCountTests = 0;
            string strTestCasesRow;

            do
            {
                strTestCasesRow = objSheet.Cells[x, 2].Text();

                if (strTestCasesRow.ToUpper() == strExecutionSet.ToUpper())
                {
                    intCountTests++;
                }
                x++;

            } while (objSheet.Cells[x, 1].Text() != "");

            objTestCases = new string[intCountTests, 6];

            x = 2;
            int y = 0;

            do
            {
                strTestCasesRow = objSheet.Cells[x, 2].Text();

                if (strTestCasesRow.ToUpper() == strExecutionSet.ToUpper())
                {
                    objTestCases[y, 0] = objSheet.Cells[x, 1].Text();
                    objTestCases[y, 1] = objSheet.Cells[x, 2].Text();
                    objTestCases[y, 2] = objSheet.Cells[x, 3].Text();
                    objTestCases[y, 3] = objSheet.Cells[x, 4].Text();
                    objTestCases[y, 4] = objSheet.Cells[x, 5].Text();
                    objTestCases[y, 5] = objSheet.Cells[x, 6].Text();
                    y++;
                }
                x++;
            } while (objSheet.Cells[x, 1].Text() != "");


            // Content
            objSheet = objWkb.Worksheets["Content"] as Worksheet;

            x = 2;
            y = 1;
            intCountX = 0;
            intCountY = 0;

            do
            {
                x++;
                intCountX++;

            } while (objSheet.Cells[x, 1].Text() != "");

            do
            {
                y++;
                intCountY++;

            } while (objSheet.Cells[1, y].Text() != "");


            objContent = new string[intCountX, intCountY];


            x = 2;
            y = 1;
            for (int r = 0; r < intCountX; r++)
            {
                for (int i = 0; i < intCountY; i++)
                {
                    objContent[r, i] = objSheet.Cells[x, y].Text();
                    y++;
                }
                y = 1;
                x++;
            }

            fnCloseExcel();
        }
        public void fnFolderSetup()
        {

            //string[] strSubFolders = new string[2] { "ScreenShots", strExecutionDate + "_" + strReportName };
            string[] strSubFolders = new string[2] { "ScreenShots", strReportName };

            bool blFExist = System.IO.Directory.Exists(strReportLocation);
            if (!blFExist)
            {
                System.IO.Directory.CreateDirectory(strReportLocation);
            }
            else
                blFExist = false;

            foreach (string strFolder in strSubFolders)
            {
                blFExist = System.IO.Directory.Exists(strReportLocation + strFolder);

                if (!blFExist)
                {
                    System.IO.Directory.CreateDirectory(strReportLocation + strFolder);
                }
            }
        }

        public void fnSSLightMethod()
        {
            //Configuration
            SLDocument objSSMethod = new SLDocument(strConfigFile, "Configuration");

            int x = 2;
            while (!string.IsNullOrEmpty(objSSMethod.GetCellValueAsString(x, 1)))
            {
                strProjectName = objSSMethod.GetCellValueAsString(x, 1);
                strReportLocation = objSSMethod.GetCellValueAsString(x, 2) + DateTime.Now.ToString("MMddyyyy_hhmmss") + @"\";
                strDataDriverLocation = objSSMethod.GetCellValueAsString(x, 3);
                strDriverSheet = objSSMethod.GetCellValueAsString(x, 4);

                x++;
            }

            string strExecutionRow;

            //Executions

            objSSMethod.SelectWorksheet("Executions");

            x = 2;

            do
            {
                strExecutionRow = objSSMethod.GetCellValueAsString(x, 5);

                if (strExecutionRow.ToUpper() == "NEW")
                {
                    strExecutionDate = objSSMethod.GetCellValueAsString(x, 1);
                    strReportName = objSSMethod.GetCellValueAsString(x, 2);
                    strReportEnv = objSSMethod.GetCellValueAsString(x, 3);
                    strExecutionSet = objSSMethod.GetCellValueAsString(x, 4);
                    strBrowser = objSSMethod.GetCellValueAsString(x, 6);

                    blExecutionFlag = true;
                    break;
                }
                else
                    blExecutionFlag = false;
                x++;
            } while (!string.IsNullOrEmpty(objSSMethod.GetCellValueAsString(x, 1)));

            //Test Cases Selection

            objSSMethod.SelectWorksheet("TestCases");

            x = 2;
            intCountTests = 0;
            string strTestCasesRow;

            do
            {
                strTestCasesRow = objSSMethod.GetCellValueAsString(x, 2);

                if (strTestCasesRow.ToUpper() == strExecutionSet.ToUpper())
                {
                    intCountTests++;
                }
                x++;

            } while (!string.IsNullOrEmpty(objSSMethod.GetCellValueAsString(x, 1)));

            objTestCases = new string[intCountTests, 6];

            x = 2;
            int y = 0;

            do
            {
                strTestCasesRow = objSSMethod.GetCellValueAsString(x, 2);

                if (strTestCasesRow.ToUpper() == strExecutionSet.ToUpper())
                {
                    objTestCases[y, 0] = objSSMethod.GetCellValueAsString(x, 1);
                    objTestCases[y, 1] = objSSMethod.GetCellValueAsString(x, 2);
                    objTestCases[y, 2] = objSSMethod.GetCellValueAsString(x, 3);
                    objTestCases[y, 3] = objSSMethod.GetCellValueAsString(x, 4);
                    objTestCases[y, 4] = objSSMethod.GetCellValueAsString(x, 5);
                    objTestCases[y, 5] = objSSMethod.GetCellValueAsString(x, 6);
                    y++;
                }
                x++;
            } while (!string.IsNullOrEmpty(objSSMethod.GetCellValueAsString(x, 1)));

            // Content
            objSSMethod.SelectWorksheet("Content");


            x = 2;
            y = 1;
            intCountX = 0;
            intCountY = 0;

            do
            {
                x++;
                intCountX++;

            } while (!string.IsNullOrEmpty(objSSMethod.GetCellValueAsString(x, 1)));

            do
            {
                y++;
                intCountY++;

            } while (!string.IsNullOrEmpty(objSSMethod.GetCellValueAsString(1, y)));


            objContent = new string[intCountX, intCountY];


            x = 2;
            y = 1;
            for (int r = 0; r < intCountX; r++)
            {
                for (int i = 0; i < intCountY; i++)
                {
                    objContent[r, i] = objSSMethod.GetCellValueAsString(x, y);
                    y++;
                }
                y = 1;
                x++;
            }
            objSSMethod.Dispose();
        }



    }
}
