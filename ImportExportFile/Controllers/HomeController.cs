using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ImportExportFile.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Configuration;
using System.Web.Mvc;
using System.Xml;

namespace ImportExportFile.Controllers
{
    public class HomeController : Controller
    {
        //
        // GET: /Home/
        Repository.Repository repo;
        Repository.ExportData export;
        Repository.ImportData import;

        public HomeController()
        {
            repo = new Repository.Repository();
            export = new Repository.ExportData();
            import = new Repository.ImportData();

        }

        public ActionResult Index()
        {
            return View();
        }

        // FILE UPLOAD 
        [HttpPost]
        public JsonResult Upload(HttpPostedFileBase file2)
        {

            HttpPostedFileBase file = Request.Files[0];
           
            if (file.ContentLength > 0)
            {
                string fileExtension = System.IO.Path.GetExtension(file.FileName);

                if (fileExtension == ".xls" || fileExtension == ".xlsx")
                {
                    string fileLocation = string.Format(@"{0}\{1}", Server.MapPath("~/App_Data/ExcelFiles"), file.FileName);

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

                    import.Import(fileLocation);

                }



            }
            
            return Json("ok", JsonRequestBehavior.AllowGet);
        }


        // GET DATA FROM TABLE
        public ActionResult getData() 
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Нефтепродукт", typeof(string));

            List<ExportList> exportData = repo.getExportData();

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

                return PartialView(dt);

            }
            else {
                return PartialView();
            }

            
            
        }

        // EXPORT DATA
        public void Export()
        {
            string filePath = string.Format("{0}/{1}", Server.MapPath("~/App_Data/ExcelFiles"), "Export.xlsx");

            export.CreateSpreadsheetWorkbook(filePath);
            
            DownloadFile(filePath);
        }


        // DOWNLOAD FILE
        public void DownloadFile(string filePath)
        {
            Response.ClearContent();
            Response.AddHeader("content-disposition", "attachment; filename=Exported_Data.xlsx");
            Response.ContentType = "application/excel";
            Response.WriteFile(filePath);
            Response.End();
        }


        

    }
}
