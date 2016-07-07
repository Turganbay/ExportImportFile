using ImportExportFile.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace ImportExportFile.Repository
{
    public class Repository
    {
        ConnectDb db;
        public Repository()
        {
            db = new ConnectDb();
        }


        //INSERT REGIONS
        public void InsertRegions(List<Region> list) { 
        
            if(db.isOpen()){
                var conn = db.GetConnection();
             
                foreach(var r in list){
           
                            SqlCommand cmd = new SqlCommand("dbo.insertRegions", conn);
                            cmd.CommandType = CommandType.StoredProcedure;

                            cmd.Parameters.AddWithValue("@name", r.name);
                            cmd.ExecuteNonQuery();
                    
                }

            }

        
        }

        //INSERT Products
        public void InsertProducts(List<Product> list)
        {

            if (db.isOpen())
            {
                var conn = db.GetConnection();

                foreach (var r in list)
                {

                    SqlCommand cmd = new SqlCommand("dbo.insertProducts", conn);
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@name", r.name);
                    cmd.ExecuteNonQuery();

                }

            }

        }

        //INSERT Company
        public void InsertCompany(List<Company> list)
        {

            if (db.isOpen())
            {
                var conn = db.GetConnection();

                foreach (var r in list)
                {

                    SqlCommand cmd = new SqlCommand("dbo.insertCompany", conn);
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@name", r.name);
                    cmd.ExecuteNonQuery();

                }

            }

        }

        //INSERT Quantity
        public void InsertQuantity(List<Quantity> list)
        {
            if (db.isOpen())
            {
                var conn = db.GetConnection();

                foreach (var r in list)
                {

                    SqlCommand cmd = new SqlCommand("dbo.insertQuantity", conn);
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@region_id", r.region_id);
                    cmd.Parameters.AddWithValue("@company_id", r.company_id);
                    cmd.Parameters.AddWithValue("@product_id", r.product_id);
                    cmd.Parameters.AddWithValue("@quantity", r.quantity);
                    cmd.Parameters.AddWithValue("@create_time", r.create_time);
    
                    cmd.ExecuteNonQuery();

                }

            }
            db.Close();

        }

        // get Region ID
        public int getRegionID(string name)
        {
            int ID = 0;

            if (db.isOpen())
            {
                    var conn = db.GetConnection();

                    SqlCommand cmd = new SqlCommand("dbo.getRegionID", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@name", name);
                    var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        ID = Convert.ToInt32(reader["id"]);
                    }

            }

            db.Close();

            return ID;

        }


        // get Company ID
        public int getCompanyID(string name)
        {
            int ID = 0;

            if (db.isOpen())
            {
                var conn = db.GetConnection();

                SqlCommand cmd = new SqlCommand("dbo.getCompanyID", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@name", name);
                var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    ID = Convert.ToInt32(reader["id"]);
                }

            }

            db.Close();


            return ID;

        }


        // get Product ID
        public int getProductID(string name)
        {
            int ID = 0;

            if (db.isOpen())
            {
                var conn = db.GetConnection();

                SqlCommand cmd = new SqlCommand("dbo.getProductID", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@name", name);
                var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    ID = Convert.ToInt32(reader["id"]);
                }

            }
            db.Close();

            return ID;
        }

        // GET Export Data
        public List<ExportList> getExportData() 
        {
            List<ExportList> list = new List<ExportList>();

            if (db.isOpen())
            {
                var conn = db.GetConnection();

                SqlCommand cmd = new SqlCommand("dbo.getExportData", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                var reader = cmd.ExecuteReader();
                
                while (reader.Read())
                {
                    ExportList e = new ExportList();
                    e.region = reader["region"].ToString();
                    e.product = reader["product"].ToString();
                    e.sum = Convert.ToDouble(reader["summa"]);
                    list.Add(e);
                }

            }

            db.Close();

            return list;
        } 

    }
}