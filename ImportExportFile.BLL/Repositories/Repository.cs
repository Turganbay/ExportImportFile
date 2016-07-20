using ImportExportFile.BLL.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using ImportExportFile.DAL.Domain;

namespace ImportExportFile.BLL.Repositories
{
    public class Repository
    {
        IEDomain db;
        public Repository()
        {
            db = new IEDomain();
        }


        //INSERT REGIONS
        public void InsertRegions(List<Region> list)
        {
            if (db.isOpen())
            {
                foreach (var r in list)
                {
                    db.InsertRegion(r.name);
                }
            }

            db.Close();
        }

        //INSERT Products
        public void InsertProducts(List<Product> list)
        {
            if (db.isOpen())
            {     
                foreach (var r in list)
                {
                    db.InsertProduct(r.name);
                }
            }
        }

        //INSERT Company
        public void InsertCompany(List<Company> list)
        {
            if (db.isOpen())
            {
                foreach (var r in list)
                {
                    db.InsertCompany(r.name);
                }
            }
        }

        //INSERT Quantity
        public void InsertQuantity(List<Quantity> list)
        {
            if (db.isOpen())
            {
                foreach (var r in list)
                {
                    db.InsertQuantity(r.region_id, r.company_id, r.product_id, r.quantity, r.create_time);
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
                ID = db.GetRegionID(name);
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
                ID = db.GetCompanyID(name);

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
                ID = db.GetProductID(name);                
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

                var reader = db.GetExportData();
                
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


        public void FileConvert(HttpPostedFileBase file, string directory)
        {
            string fileExtension = System.IO.Path.GetExtension(file.FileName);
            string fileLocation = string.Format(@"{0}\{1}", directory, file.FileName);

            if (fileExtension == ".xls" || fileExtension == ".xlsx")
            {
                //string fileLocation = string.Format(@"{0}\{1}", Server.MapPath("~/App_Data/ExcelFiles"), file.FileName);

                if (System.IO.File.Exists(fileLocation))
                {
                    System.IO.File.Delete(fileLocation);
                    System.IO.File.Delete(fileLocation + "x");
                }

                file.SaveAs(fileLocation);

                if (fileExtension == ".xls")
                {

                    var app = new Microsoft.Office.Interop.Excel.Application();
                    var wb = app.Workbooks.Open(fileLocation);

                    fileLocation += "x";

                    wb.SaveAs(Filename: fileLocation, FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                    wb.Close();
                    app.Quit();
                }

                ImportData c = new ImportData();
                c.Import(fileLocation);

            }
        }

        // get Data Table
        public DataTable GetDataTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Нефтепродукт", typeof(string));

            List<ExportList> exportData = getExportData();

            if (exportData.Count != 0)
            {
                List<string> r = new List<string>();
                List<string> p = new List<string>();

                Dictionary<string, string> dict = new Dictionary<string, string>();


                foreach (var item in exportData)
                {
                    if (!r.Contains(item.region))
                    {
                        r.Add(item.region);
                        dt.Columns.Add(item.region, typeof(string));

                    }
                    if (!p.Contains(item.product))
                    {
                        p.Add(item.product);
                        //dt.Rows.Add(item.product);
                        dict[item.product] = "";
                    }

                    dict[item.product] += item.sum + "|";


                }

                foreach (string name in p)
                {
                    int i = 0;
                    DataRow newRow = dt.NewRow();
                    string str = dict[name];
                    char delimiterChar = '|';

                    newRow[i] = name;
                    string[] lines = str.Split(delimiterChar);

                    foreach (string line in lines)
                    {
                        i++;
                        if (!String.IsNullOrEmpty(line))
                        {
                            newRow[i] = line;
                        }
                    }

                    dt.Rows.Add(newRow);
                }

               
            }

            return dt;

        }

    }
}