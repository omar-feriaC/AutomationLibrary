using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationLibrary
{
    public class ClsDB
    {
        private OracleConnection conn;
        private static string strConnection;

        /// <summary>
        /// Returns the string connection for based on the parameters provided
        /// </summary>
        /// <param name="strHost"></param>
        /// <param name="strPort"></param>
        /// <param name="strService"></param>
        /// <param name="strUser"></param>
        /// <param name="strPassword"></param>
        /// <returns></returns>
        public string GetConnectionString(string strHost, string strPort, string strService, string strUser, string strPassword)
        {
            strConnection = "Data Source=" +
                            "(DESCRIPTION =" + "" +
                                "(ADDRESS = " +
                                    "(PROTOCOL = TCP)" +
                                    "(HOST = " + strHost + ")" +
                                    "(PORT = " + strPort + "))" +
                            "(CONNECT_DATA = " +
                                "(SERVER = DEDICATED)" +
                                "(SERVICE_NAME = " + strService + ")));" +
                            "Password=" + strPassword + ";User ID=" + strUser + "";
            return strConnection;
        }

        /// <summary>
        /// Opens a connection for DB
        /// </summary>
        /// <param name="pstrConnection"></param>
        public void fnOpenConnection(string pstrConnection)
        {
            try
            {
                conn = new OracleConnection(pstrConnection);
                conn.Open();
            }
            catch (Exception e)
            {
                //Stack Trace
                ClsVariables.TempStackTrace = e.StackTrace;
                ClsVariables.fnAddStackTrace(ClsVariables.TempStackTrace);
                Console.WriteLine("The connection cannot be opened: Error -> " + e.Message);
            }
        }

        /// <summary>
        /// Close a connection DB
        /// </summary>
        public void fnCloseConnection()
        {
            try
            {
                conn.Dispose();
                conn.Close();
            }
            catch (Exception e)
            {
                //Stack Trace
                ClsVariables.TempStackTrace = e.StackTrace;
                ClsVariables.fnAddStackTrace(ClsVariables.TempStackTrace);
                Console.WriteLine("The connection cannot be closed: Error -> " + e.Message);
            }
        }

        /// <summary>
        /// Executes a query and  close the connection
        /// </summary>
        /// <param name="pstrQuery"></param>
        public void fnExecuteQuery(string pstrQuery)
        {
            try
            {
                OracleCommand cmd = new OracleCommand(pstrQuery, conn);
                cmd.ExecuteNonQuery();
                fnCloseConnection();
            }
            catch (Exception e)
            {
                //Stack Trace
                ClsVariables.TempStackTrace = e.StackTrace;
                ClsVariables.fnAddStackTrace(ClsVariables.TempStackTrace);
                fnCloseConnection();
                Console.WriteLine("The query cannot be executed: Error -> " + e.Message);
            }
        }

        /// <summary>
        /// Returns a data reader with teh result of a query
        /// </summary>
        /// <param name="pstrQuery"></param>
        /// <returns></returns>
        public DataTable fnDataReader(string pstrQuery)
        {
            try
            {
                OracleCommand cmd = new OracleCommand(pstrQuery, conn);
                DbDataReader reader = cmd.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(reader);
                fnCloseConnection();
                return dt;
            }
            catch (Exception e)
            {
                //Stack Trace
                ClsVariables.TempStackTrace = e.StackTrace;
                ClsVariables.fnAddStackTrace(ClsVariables.TempStackTrace);
                fnCloseConnection();
                Console.WriteLine("The query cannot be executed: Error -> " + e.Message);
                return null;
            }
        }

        /// <summary>
        /// Returns a datatable with the results of a query
        /// </summary>
        /// <param name="pstrQuery"></param>
        /// <returns></returns>
        public DataTable fnDataSet(string pstrQuery)
        {
            try
            {
                OracleDataAdapter adapter = new OracleDataAdapter(pstrQuery, conn);
                OracleCommandBuilder builder = new OracleCommandBuilder(adapter);
                DataSet dataset = new DataSet();
                adapter.Fill(dataset);
                DataTable datatable = new DataTable();
                datatable = dataset.Tables[0];
                fnCloseConnection();
                return datatable;
            }
            catch (Exception e)
            {
                //Stack Trace
                ClsVariables.TempStackTrace = e.StackTrace;
                ClsVariables.fnAddStackTrace(ClsVariables.TempStackTrace);
                fnCloseConnection();
                Console.WriteLine("The data table cannot be created: Error -> " + e.Message);
                return null;
            }
        }

        /// <summary>
        /// Returns only one value of a query
        /// </summary>
        /// <param name="pstrQuery"></param>
        /// <returns></returns>
        public String fnGetSingleValue(string pstrQuery)
        {
            string strValue;
            try
            {
                using (var cmd = new OracleCommand(pstrQuery, conn))
                {
                    strValue = cmd.ExecuteScalar().ToString();
                    fnCloseConnection();
                }
            }
            catch (Exception e)
            {
                //Stack Trace
                ClsVariables.TempStackTrace = e.StackTrace;
                ClsVariables.fnAddStackTrace(ClsVariables.TempStackTrace);
                fnCloseConnection();
                Console.WriteLine("The data table cannot be created: Error -> " + e.Message);
                strValue = "";
            }

            return strValue;
        }

    }
}
