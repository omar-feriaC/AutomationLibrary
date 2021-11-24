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
    public class clsDB
    {
        private OracleConnection conn;
        private static string strConnection;

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

        public void fnOpenConnection(string pstrConnection)
        {
            try
            {
                conn = new OracleConnection(pstrConnection);
                conn.Open();
            }
            catch (Exception e)
            {
                Console.WriteLine("The connection cannot be opened: Error -> " + e.Message);
            }
        }

        public void fnCloseConnection()
        {
            try
            {
                conn.Dispose();
                conn.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("The connection cannot be closed: Error -> " + e.Message);
            }
        }

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
                fnCloseConnection();
                Console.WriteLine("The query cannot be executed: Error -> " + e.Message);
            }
        }

        public DbDataReader fnDataReader(string pstrQuery)
        {
            try
            {
                OracleCommand cmd = new OracleCommand(pstrQuery, conn);
                DbDataReader reader = cmd.ExecuteReader();
                fnCloseConnection();
                return reader;
            }
            catch (Exception e)
            {
                fnCloseConnection();
                Console.WriteLine("The query cannot be executed: Error -> " + e.Message);
                return null;
            }
        }

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
                fnCloseConnection();
                Console.WriteLine("The data table cannot be created: Error -> " + e.Message);
                return null;
            }
        }

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
                fnCloseConnection();
                Console.WriteLine("The data table cannot be created: Error -> " + e.Message);
                strValue = "";
            }

            return strValue;
        }

    }
}
