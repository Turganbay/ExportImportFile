using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Configuration;

namespace ImportExportFile.DAL.Domain
{
    public class IEDomain
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


        public IEDomain()
        {
            con = GetConnection();
        }


        // insert Region
        public void InsertRegion(string name)
        {

            SqlCommand cmd = new SqlCommand("dbo.insertRegions", con);
            cmd.CommandType = CommandType.StoredProcedure;

            cmd.Parameters.AddWithValue("@name", name);
            cmd.ExecuteNonQuery();
        }

        // insert Product
        public void InsertProduct(string name)
        {
            SqlCommand cmd = new SqlCommand("dbo.insertProducts", con);
            cmd.CommandType = CommandType.StoredProcedure;

            cmd.Parameters.AddWithValue("@name", name);
            cmd.ExecuteNonQuery();
        }

        // insert Company
        public void InsertCompany(string name)
        {
            SqlCommand cmd = new SqlCommand("dbo.insertCompany", con);
            cmd.CommandType = CommandType.StoredProcedure;

            cmd.Parameters.AddWithValue("@name", name);
            cmd.ExecuteNonQuery();
        }

        // insert Quantity
        public void InsertQuantity(int region_id, int company_id, int product_id, double quantity, DateTime create_time)
        {
            SqlCommand cmd = new SqlCommand("dbo.insertQuantity", con);
            cmd.CommandType = CommandType.StoredProcedure;

            cmd.Parameters.AddWithValue("@region_id", region_id);
            cmd.Parameters.AddWithValue("@company_id", company_id);
            cmd.Parameters.AddWithValue("@product_id", product_id);
            cmd.Parameters.AddWithValue("@quantity", quantity);
            cmd.Parameters.AddWithValue("@create_time", create_time);

            cmd.ExecuteNonQuery();
        }

        // get Region Id
        public int GetRegionID(string name)
        {
            int ID = 0;
            SqlCommand cmd = new SqlCommand("dbo.getRegionID", con);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@name", name);
            var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                ID = Convert.ToInt32(reader["id"]);
            }

            return ID;
        }

        public int GetCompanyID(string name)
        {
            int ID = 0;
            SqlCommand cmd = new SqlCommand("dbo.getCompanyID", con);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@name", name);
            var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                ID = Convert.ToInt32(reader["id"]);
            }
            return ID;
        }

        //get product id
        public int GetProductID(string name)
        {
            int ID = 0;
            SqlCommand cmd = new SqlCommand("dbo.getProductID", con);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@name", name);
            var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                ID = Convert.ToInt32(reader["id"]);
            }
            return ID;
        }

        // get export data
        public SqlDataReader GetExportData()
        {
            SqlCommand cmd = new SqlCommand("dbo.getExportData", con);
            cmd.CommandType = CommandType.StoredProcedure;
            var reader = cmd.ExecuteReader();

            return reader;
        }
    }
}