using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Configuration;

namespace ImportExportFile.Repository
{
    public class ConnectDb
    {

        protected SqlConnection con;

        // _CONNECTION_OPEN
        public bool isOpen(string Connection = "DefaultConnection")
        {
            con = new SqlConnection(@WebConfigurationManager.ConnectionStrings[Connection].ToString());

            try
            {
                bool b = true;
                if (con.State.ToString() != "Open")
                {
                    con.Open();
                }
                return b;
            }
            catch (SqlException ex)
            {
                return false;
            }
        }

        // _CONNECTION_STRING
        public SqlConnection GetConnection()
        {
            return con;
        }

        // _CONNECTION_CLOSE
        public bool Close()
        {
            try
            {
                con.Close();
                return false;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

    }
}